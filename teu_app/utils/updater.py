"""
updater.py
----------
Mise à jour mensuelle des RATIOS L/D à partir d'un fichier Excel de moves.

Colonnes attendues dans le fichier de moves :
  - Lane              : identifiant de la lane
  - D/L               : 'L' = Load, 'D' = Discharge  (pas D/L/S)
  - Dis./Load/AS Completed Date : date du mouvement

Logique :
  1. Lire le fichier Excel uploadé
  2. Détecter les colonnes Lane, D/L, et date
  3. Filtrer les 3 derniers mois (depuis la date max dans le fichier)
  4. Calculer les ratios par lane
  5. Merger avec ratios.json existant et sauvegarder
"""

import os
import json
import io
from datetime import datetime, timedelta

import pandas as pd

# Chemin vers le fichier JSON de persistance
DATA_DIR   = os.path.join(os.path.dirname(__file__), "..", "data")
RATIOS_PATH = os.path.join(DATA_DIR, "ratios.json")

# Fichier de log de la dernière mise à jour
UPDATE_LOG_PATH = os.path.join(DATA_DIR, "last_update.json")

# Nombre de mois d'historique à conserver pour le calcul des ratios
MONTHS_WINDOW = 3


# ── Chargement / Sauvegarde JSON ───────────────────────────────────────────────

def load_ratios() -> dict:
    """Charge les RATIOS depuis le fichier JSON."""
    if not os.path.exists(RATIOS_PATH):
        return {}
    with open(RATIOS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_ratios(ratios: dict) -> None:
    """Sauvegarde les RATIOS dans le fichier JSON."""
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(RATIOS_PATH, "w", encoding="utf-8") as f:
        json.dump(ratios, f, ensure_ascii=False, indent=2)


def load_update_log() -> dict:
    """Charge le log de la dernière mise à jour."""
    if not os.path.exists(UPDATE_LOG_PATH):
        return {}
    with open(UPDATE_LOG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def save_update_log(log: dict) -> None:
    """Sauvegarde le log de mise à jour."""
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(UPDATE_LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)


# ── Détection de colonnes ──────────────────────────────────────────────────────

def _find_col(columns: list[str], keywords: list[str]) -> str | None:
    """
    Cherche dans une liste de colonnes (lowercase) celle qui contient
    au moins un des mots-clés.
    """
    for col in columns:
        for kw in keywords:
            if kw in col:
                return col
    return None


# ── Calcul des ratios depuis un DataFrame ──────────────────────────────────────

def _compute_ratios_from_df(df: pd.DataFrame,
                             lane_col: str,
                             dl_col: str,
                             date_col: str) -> dict:
    """
    Calcule les ratios Load/Discharge par lane sur les MONTHS_WINDOW derniers mois.
    Retourne un dict {lane: {...}} au format RATIOS.
    """
    # Convertir les dates
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=False, errors="coerce")
    df = df.dropna(subset=[date_col, lane_col, dl_col])

    # Filtrer uniquement L ou D (exclure S — shifts)
    df[dl_col] = df[dl_col].astype(str).str.strip().str.upper()
    df = df[df[dl_col].isin(["L", "D"])]

    if df.empty:
        return {}

    # Déterminer la date max et la fenêtre de 3 mois
    max_date   = df[date_col].max()
    start_date = max_date - pd.DateOffset(months=MONTHS_WINDOW)

    df_window = df[df[date_col] >= start_date].copy()
    if df_window.empty:
        return {}

    # Normaliser les lanes en majuscules
    df_window[lane_col] = df_window[lane_col].astype(str).str.strip().str.upper()

    new_ratios = {}
    for lane, grp in df_window.groupby(lane_col):
        if not lane or lane in ("NAN", ""):
            continue

        total_L = int((grp[dl_col] == "L").sum())
        total_D = int((grp[dl_col] == "D").sum())
        total   = total_L + total_D

        if total == 0:
            continue

        pct_L = round(total_L / total * 100, 2)
        pct_D = round(total_D / total * 100, 2)

        # n_voyages = nombre de dates distinctes (jours d'escale)
        n_voyages = int(grp[date_col].dt.date.nunique())

        avg_L = round(total_L / n_voyages, 1) if n_voyages else 0.0
        avg_D = round(total_D / n_voyages, 1) if n_voyages else 0.0

        # last_date et period_start au format YYYY-MM-DD
        lane_dates  = grp[date_col]
        last_date   = lane_dates.max().strftime("%Y-%m-%d")
        period_start = lane_dates.min().strftime("%Y-%m-%d")

        new_ratios[lane] = {
            "last_date":    last_date,
            "period_start": period_start,
            "total_L":      total_L,
            "total_D":      total_D,
            "total":        total,
            "pct_L":        pct_L,
            "pct_D":        pct_D,
            "n_voyages":    n_voyages,
            "avg_L":        avg_L,
            "avg_D":        avg_D,
        }

    return new_ratios


# ── Fonction principale ────────────────────────────────────────────────────────

def process_moves_file(file_bytes: bytes, filename: str) -> dict:
    """
    Point d'entrée principal.
    Lit un fichier de moves Excel/CSV, recalcule les ratios,
    met à jour ratios.json et retourne un rapport de la mise à jour.

    Retourne :
        {
            "success": bool,
            "error": str | None,
            "updated_lanes": [...],   # lanes existantes mises à jour
            "new_lanes": [...],       # nouvelles lanes ajoutées
            "unchanged_lanes": [...], # lanes non présentes dans le fichier
            "total_rows": int,        # lignes valides traitées
            "period_start": str,
            "period_end": str,
        }
    """
    # ── 1. Lire le fichier ────────────────────────────────────────────────────
    try:
        if filename.lower().endswith(".csv"):
            df = pd.read_csv(io.BytesIO(file_bytes))
        else:
            df = pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        return {"success": False, "error": f"Impossible de lire le fichier : {e}"}

    # ── 2. Normaliser les noms de colonnes ────────────────────────────────────
    original_cols = list(df.columns)
    df.columns    = [str(c).strip().lower() for c in df.columns]
    col_map       = dict(zip(df.columns, original_cols))  # lower → original

    # ── 3. Détecter les colonnes utiles ───────────────────────────────────────
    lane_col = _find_col(df.columns, ["lane"])
    dl_col   = _find_col(df.columns, ["d/l"])          # cherche exactement "d/l"
    date_col = _find_col(df.columns, [
        "dis./load/as completed date",
        "completed date",
        "as completed",
        "move date",
        "date",
    ])

    # Vérification stricte de D/L : on veut "d/l" pas "d/l/s"
    # Si plusieurs colonnes matchent, prendre la plus courte (= exacte)
    dl_candidates = [c for c in df.columns if "d/l" in c]
    if dl_candidates:
        dl_col = min(dl_candidates, key=len)   # "d/l" < "d/l/s"

    if lane_col is None:
        return {"success": False, "error": "Colonne 'Lane' introuvable dans le fichier."}
    if dl_col is None:
        return {"success": False, "error": "Colonne 'D/L' introuvable dans le fichier."}
    if date_col is None:
        return {"success": False, "error": (
            "Colonne de date introuvable. "
            "Attendu : 'Dis./Load/AS Completed Date' ou similaire."
        )}

    # ── 4. Calculer les nouveaux ratios ───────────────────────────────────────
    new_ratios = _compute_ratios_from_df(df, lane_col, dl_col, date_col)

    if not new_ratios:
        return {
            "success": False,
            "error": "Aucune donnée valide trouvée (vérifiez les colonnes D/L et date)."
        }

    # ── 5. Merger avec les RATIOS existants ───────────────────────────────────
    existing = load_ratios()
    existing_keys = set(existing.keys())
    new_keys      = set(new_ratios.keys())

    updated_lanes   = sorted(existing_keys & new_keys)        # intersection
    new_lanes       = sorted(new_keys - existing_keys)        # nouvelles
    unchanged_lanes = sorted(existing_keys - new_keys)        # non touchées

    merged = {**existing, **new_ratios}   # new_ratios écrase existing pour les lanes communes
    save_ratios(merged)

    # ── 6. Log de mise à jour ──────────────────────────────────────────────────
    all_dates = df[date_col]
    all_dates = pd.to_datetime(all_dates, errors="coerce").dropna()
    max_date  = all_dates.max()
    period_end   = max_date.strftime("%Y-%m-%d") if not pd.isnull(max_date) else "?"
    period_start_dt = max_date - pd.DateOffset(months=MONTHS_WINDOW)
    period_start = period_start_dt.strftime("%Y-%m-%d") if not pd.isnull(max_date) else "?"

    log = {
        "updated_at":     datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "filename":       filename,
        "total_rows":     len(df),
        "updated_lanes":  updated_lanes,
        "new_lanes":      new_lanes,
        "unchanged_lanes": unchanged_lanes,
        "period_start":   period_start,
        "period_end":     period_end,
    }
    save_update_log(log)

    return {
        "success":          True,
        "error":            None,
        "updated_lanes":    updated_lanes,
        "new_lanes":        new_lanes,
        "unchanged_lanes":  unchanged_lanes,
        "total_rows":       len(df),
        "period_start":     period_start,
        "period_end":       period_end,
    }
