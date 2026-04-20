"""
fm_predictor.py
---------------
Logique de prédiction Full / Empty appliquée sur les résultats Load/Discharge.

Pipeline en 2 étapes :
  Étape 1 (predictor.py) : Lane + Volume  →  pred_L  +  pred_D
  Étape 2 (ce fichier)   : pred_L         →  L_Full  +  L_Empty
                           pred_D         →  D_Full  +  D_Empty

Les ratios F/M sont calculés séparément pour D et L,
sur les 3 derniers mois avant la dernière escale de chaque lane.

Fichier de données : data/fm_ratios.json  (mis à jour via updater_fm.py)
"""

import os
import json
import io
import pandas as pd
from datetime import datetime

_FM_RATIOS_PATH = os.path.join(os.path.dirname(__file__), '..', 'data', 'fm_ratios.json')


def _load_fm_ratios() -> dict:
    try:
        with open(_FM_RATIOS_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except Exception as e:
        print(f'[fm_predictor] Erreur chargement fm_ratios.json : {e}')
        return {}


FM_RATIOS: dict = _load_fm_ratios()


def reload_fm_ratios() -> int:
    global FM_RATIOS
    FM_RATIOS = _load_fm_ratios()
    return len(FM_RATIOS)


def predict_fm(lane: str, pred_L: int, pred_D: int) -> dict | None:
    """
    Applique les ratios F/M sur pred_L et pred_D.
    Retourne None si la lane est inconnue dans FM_RATIOS.
    """
    r = FM_RATIOS.get(lane.upper())
    if r is None:
        return None

    rd = r.get('discharge') or {}
    rl = r.get('load')      or {}

    pct_D_full  = rd.get('pct_F', 50.0)
    pct_D_empty = rd.get('pct_M', 50.0)
    pct_L_full  = rl.get('pct_F', 50.0)
    pct_L_empty = rl.get('pct_M', 50.0)

    return {
        'L_full':      round(pred_L * pct_L_full  / 100),
        'L_empty':     round(pred_L * pct_L_empty / 100),
        'D_full':      round(pred_D * pct_D_full  / 100),
        'D_empty':     round(pred_D * pct_D_empty / 100),
        'pct_L_full':  pct_L_full,
        'pct_L_empty': pct_L_empty,
        'pct_D_full':  pct_D_full,
        'pct_D_empty': pct_D_empty,
        'fm_period_start': r.get('period_start', ''),
        'fm_last_date':    r.get('last_date', ''),
        'fm_available': True,
    }


def predict_full_pipeline(lane: str, volume: int) -> dict | None:
    """
    Pipeline complet : Lane + Volume → pred_L/D → L_Full, L_Empty, D_Full, D_Empty.
    Retourne None si la lane est inconnue dans les ratios D/L.
    """
    from utils.predictor import predict_single
    step1 = predict_single(lane, volume)
    if step1 is None:
        return None

    step2 = predict_fm(lane, step1['pred_L'], step1['pred_D'])
    result = {**step1}
    if step2:
        result.update(step2)
    else:
        result.update({
            'L_full': None, 'L_empty': None,
            'D_full': None, 'D_empty': None,
            'pct_L_full': None, 'pct_L_empty': None,
            'pct_D_full': None, 'pct_D_empty': None,
            'fm_available': False,
        })
    return result


def predict_fm_batch(file_bytes: bytes, filename: str) -> tuple[list, list]:
    """Pipeline complet batch depuis un fichier Excel/CSV (Lane + Volume)."""
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
        res = predict_full_pipeline(lane, volume)
        if res is None:
            res = {'lane': lane, 'volume': volume,
                   'pred_L': None, 'pred_D': None, 'unknown': True}
        results.append(res)

    return results, errors
