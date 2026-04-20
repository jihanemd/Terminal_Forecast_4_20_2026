"""
teu_predictor.py
----------------
Étape 3 du pipeline : calcule les TEU à partir des résultats Full/Empty.

Logique TEU :
  - Chaque conteneur a un type : 1 (20 pieds) ou 2 (40 pieds)
  - TEU = count(type=1) × 1 + count(type=2) × 2
  - Exemple : 3 conteneurs de 20' + 2 de 40' → TEU = 3 + 4 = 7

Pipeline complet :
  Lane + Volume
    → pred_L, pred_D                       (predictor.py   — ratios D/L)
    → L_Full, L_Empty, D_Full, D_Empty     (fm_predictor.py — ratios F/M)
    → L_Full_TEU, L_Empty_TEU              (ce fichier     — ratios TEU par F/M et D/L)
       D_Full_TEU, D_Empty_TEU

Structure de teu_ratios.json par lane :
  {
    "FAL1WB": {
      "last_date": "...", "period_start": "...",
      "discharge": {
        "full":  { "containers": N, "teu_20": N, "teu_40": N, "teu_total": N, "pct_teu": % },
        "empty": { ... },
        "total_teu": N
      },
      "load": { "full": {...}, "empty": {...}, "total_teu": N }
    }
  }

Les ratios TEU_per_container sont dérivés de ces données :
  teu_per_container_full  = teu_total_full  / containers_full
  teu_per_container_empty = teu_total_empty / containers_empty
On applique ensuite ces ratios sur pred_Full et pred_Empty (nombre de containers).
"""

import os
import json
import io
import pandas as pd
from datetime import datetime

_TEU_RATIOS_PATH = os.path.join(os.path.dirname(__file__), '..', 'data', 'teu_ratios.json')


# ── Chargement ─────────────────────────────────────────────────────────────────

def _load_teu_ratios() -> dict:
    try:
        with open(_TEU_RATIOS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except Exception as e:
        print(f'[teu_predictor] Erreur chargement teu_ratios.json : {e}')
        return {}


TEU_RATIOS: dict = _load_teu_ratios()


def reload_teu_ratios() -> int:
    global TEU_RATIOS
    TEU_RATIOS = _load_teu_ratios()
    return len(TEU_RATIOS)


# ── Calcul TEU ─────────────────────────────────────────────────────────────────

def _teu_per_container(data: dict) -> float:
    """
    Calcule le ratio moyen TEU / container pour un groupe (ex: D-Full).
    Si pas de containers, retourne 1.5 (valeur neutre entre 1 et 2).
    """
    containers = data.get('containers', 0)
    teu_total  = data.get('teu_total', 0)
    if containers == 0:
        return 1.5
    return teu_total / containers


def predict_teu(lane: str, L_full: int, L_empty: int, D_full: int, D_empty: int) -> dict | None:
    """
    Étape 3 du pipeline : applique les ratios TEU sur les quantités F/M.

    Paramètres :
        lane    : nom de la lane
        L_full  : nombre de containers Load Full  (issu de fm_predictor.py)
        L_empty : nombre de containers Load Empty
        D_full  : nombre de containers Discharge Full
        D_empty : nombre de containers Discharge Empty

    Retourne :
        {
            'L_full_teu' : int,   TEU pour Load Full
            'L_empty_teu': int,   TEU pour Load Empty
            'D_full_teu' : int,   TEU pour Discharge Full
            'D_empty_teu': int,   TEU pour Discharge Empty
            'L_total_teu': int,   TEU Load total
            'D_total_teu': int,   TEU Discharge total
            'total_teu'  : int,   TEU global
            'teu_available': True
        }
    Retourne None si la lane est inconnue dans TEU_RATIOS.
    """
    r = TEU_RATIOS.get(lane.upper())
    if r is None:
        return None

    rd = r.get('discharge', {})
    rl = r.get('load', {})

    # Ratios TEU/container pour chaque groupe
    ratio_D_full  = _teu_per_container(rd.get('full',  {}))
    ratio_D_empty = _teu_per_container(rd.get('empty', {}))
    ratio_L_full  = _teu_per_container(rl.get('full',  {}))
    ratio_L_empty = _teu_per_container(rl.get('empty', {}))

    L_full_teu  = round((L_full  or 0) * ratio_L_full)
    L_empty_teu = round((L_empty or 0) * ratio_L_empty)
    D_full_teu  = round((D_full  or 0) * ratio_D_full)
    D_empty_teu = round((D_empty or 0) * ratio_D_empty)

    L_total_teu = L_full_teu + L_empty_teu
    D_total_teu = D_full_teu + D_empty_teu

    return {
        'L_full_teu':   L_full_teu,
        'L_empty_teu':  L_empty_teu,
        'D_full_teu':   D_full_teu,
        'D_empty_teu':  D_empty_teu,
        'L_total_teu':  L_total_teu,
        'D_total_teu':  D_total_teu,
        'total_teu':    L_total_teu + D_total_teu,
        # Ratios TEU/container (pour transparence)
        'ratio_D_full':  round(ratio_D_full, 3),
        'ratio_D_empty': round(ratio_D_empty, 3),
        'ratio_L_full':  round(ratio_L_full, 3),
        'ratio_L_empty': round(ratio_L_empty, 3),
        'teu_period_start': r.get('period_start', ''),
        'teu_last_date':    r.get('last_date', ''),
        'teu_available': True,
    }


# ── Pipeline complet ────────────────────────────────────────────────────────────

def predict_full_pipeline_with_teu(lane: str, volume: int) -> dict | None:
    """
    Pipeline complet en 3 étapes :
      Lane + Volume
        → pred_L, pred_D         (predictor.py)
        → L_Full, L_Empty, ...   (fm_predictor.py)
        → L_Full_TEU, ...        (teu_predictor.py — ce fichier)

    Retourne None si la lane est inconnue dans les ratios D/L.
    """
    from utils.fm_predictor import predict_full_pipeline

    result = predict_full_pipeline(lane, volume)
    if result is None:
        return None

    if result.get('fm_available'):
        teu = predict_teu(
            lane,
            result.get('L_full', 0),
            result.get('L_empty', 0),
            result.get('D_full', 0),
            result.get('D_empty', 0),
        )
        if teu:
            result.update(teu)
        else:
            result.update({
                'L_full_teu': None, 'L_empty_teu': None,
                'D_full_teu': None, 'D_empty_teu': None,
                'L_total_teu': None, 'D_total_teu': None,
                'total_teu': None,
                'teu_available': False,
            })
    else:
        result.update({
            'L_full_teu': None, 'L_empty_teu': None,
            'D_full_teu': None, 'D_empty_teu': None,
            'L_total_teu': None, 'D_total_teu': None,
            'total_teu': None,
            'teu_available': False,
        })

    return result


def predict_teu_batch(file_bytes: bytes, filename: str) -> tuple[list, list]:
    """Pipeline complet 3 étapes pour un fichier batch (Lane + Volume)."""
    results = []
    errors  = []

    try:
        df = pd.read_csv(io.BytesIO(file_bytes)) if filename.endswith('.csv') \
             else pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        return [], [f'Impossible de lire le fichier : {e}']

    df.columns = [str(c).strip().lower() for c in df.columns]
    lane_col = next((c for c in df.columns if 'lane' in c), None)
    vol_col  = next((c for c in df.columns
                     if any(k in c for k in ['vol', 'move', 'teu', 'total'])), None)

    if lane_col is None:
        return [], ["Colonne 'Lane' introuvable dans le fichier."]
    if vol_col is None:
        return [], ["Colonne 'Volume' introuvable dans le fichier."]

    for idx, row in df.iterrows():
        lane = str(row[lane_col]).strip().upper()
        try:
            volume = int(float(row[vol_col]))
        except (ValueError, TypeError):
            errors.append(f'Ligne {idx+2} : volume invalide ({row[vol_col]})')
            continue
        if not lane or lane in ('NAN', ''):
            errors.append(f'Ligne {idx+2} : lane vide')
            continue
        res = predict_full_pipeline_with_teu(lane, volume)
        if res is None:
            res = {'lane': lane, 'volume': volume, 'unknown': True}
        results.append(res)

    return results, errors
