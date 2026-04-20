"""
predictor.py
------------
Logique de prédiction Load / Discharge par lane.
Basé sur les ratios calculés sur les 3 derniers mois d'activité de chaque lane.

Les RATIOS sont chargés depuis data/ratios.json (mis à jour via upload mensuel).
"""

import os
import json
import pandas as pd
import io
from datetime import datetime

# ── Données de référence ───────────────────────────────────────────────────────
# Les RATIOS sont chargés depuis data/ratios.json.
# Ce fichier est mis à jour automatiquement via l'upload mensuel de moves.
# Pour chaque lane : last_date, period_start, total_L, total_D, total,
#                    pct_L, pct_D, n_voyages, avg_L, avg_D

_RATIOS_PATH = os.path.join(os.path.dirname(__file__), "..", "data", "ratios.json")


def _load_ratios_from_json() -> dict:
    """Charge les RATIOS depuis data/ratios.json."""
    try:
        with open(_RATIOS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except Exception as e:
        print(f"[predictor] Erreur chargement ratios.json : {e}")
        return {}


# Chargement initial au démarrage de l'application
RATIOS: dict = _load_ratios_from_json()


def reload_ratios() -> int:
    """
    Recharge les RATIOS depuis ratios.json (après une mise à jour mensuelle).
    Retourne le nombre de lanes chargées.
    """
    global RATIOS
    RATIOS = _load_ratios_from_json()
    return len(RATIOS)


def _is_outdated(last_date_str: str) -> bool:
    """Retourne True si la dernière escale est avant 2025."""
    return datetime.strptime(last_date_str, "%Y-%m-%d") < datetime(2025, 1, 1)


def predict_single(lane: str, volume: int) -> dict | None:
    """
    Prédit Load et Discharge pour une lane et un volume donnés.
    Retourne None si la lane est inconnue.
    """
    r = RATIOS.get(lane.upper())
    if r is None:
        return None

    pred_L = round(volume * r["pct_L"] / 100)
    pred_D = round(volume * r["pct_D"] / 100)

    return {
        "lane":         lane.upper(),
        "volume":       volume,
        "pred_L":       pred_L,
        "pred_D":       pred_D,
        "pct_L":        r["pct_L"],
        "pct_D":        r["pct_D"],
        "period_start": r["period_start"],
        "last_date":    r["last_date"],
        "n_voyages":    r["n_voyages"],
        "avg_L":        r["avg_L"],
        "avg_D":        r["avg_D"],
        "outdated":     _is_outdated(r["last_date"]),
    }


def predict_batch(file_bytes: bytes, filename: str) -> tuple[list, list]:
    """
    Prédit Load et Discharge pour chaque ligne d'un fichier Excel ou CSV.
    Retourne (results, errors).
    """
    results = []
    errors  = []

    try:
        if filename.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(file_bytes))
        else:
            df = pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        return [], [f"Impossible de lire le fichier : {e}"]

    # Normaliser les noms de colonnes
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Trouver la colonne Lane
    lane_col = next((c for c in df.columns if "lane" in c), None)
    # Trouver la colonne Volume
    vol_col  = next((c for c in df.columns if any(k in c for k in ["vol", "move", "teu", "total"])), None)

    if lane_col is None:
        return [], ["Colonne 'Lane' introuvable dans le fichier."]
    if vol_col is None:
        return [], ["Colonne 'Volume' introuvable dans le fichier."]

    for idx, row in df.iterrows():
        lane   = str(row[lane_col]).strip().upper()
        try:
            volume = int(float(row[vol_col]))
        except (ValueError, TypeError):
            errors.append(f"Ligne {idx+2} : volume invalide ({row[vol_col]})")
            continue

        if not lane or lane in ("NAN", ""):
            errors.append(f"Ligne {idx+2} : lane vide")
            continue

        res = predict_single(lane, volume)
        if res is None:
            res = {
                "lane":    lane,
                "volume":  volume,
                "pred_L":  None,
                "pred_D":  None,
                "unknown": True,
            }
        results.append(res)

    return results, errors
