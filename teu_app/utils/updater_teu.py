"""
updater_teu.py
--------------
Mise à jour mensuelle des ratios TEU depuis un fichier de moves.

Colonnes attendues : Lane · D/L/S (ou D/L) · F/M · TEU · date
Fichier de sortie  : data/teu_ratios.json
Log                : data/last_update_teu.json

Logique TEU :
  - TEU=1 → conteneur 20 pieds → contribue 1 TEU
  - TEU=2 → conteneur 40 pieds → contribue 2 TEU
  - TEU total groupe = count(TEU=1)×1 + count(TEU=2)×2
"""

import os, json, io
from datetime import datetime
import pandas as pd

DATA_DIR            = os.path.join(os.path.dirname(__file__), '..', 'data')
TEU_RATIOS_PATH     = os.path.join(DATA_DIR, 'teu_ratios.json')
TEU_UPDATE_LOG_PATH = os.path.join(DATA_DIR, 'last_update_teu.json')
MONTHS_WINDOW       = 3


def load_teu_ratios() -> dict:
    if not os.path.exists(TEU_RATIOS_PATH):
        return {}
    with open(TEU_RATIOS_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_teu_ratios(ratios: dict) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(TEU_RATIOS_PATH, 'w', encoding='utf-8') as f:
        json.dump(ratios, f, ensure_ascii=False, indent=2)

def load_teu_update_log() -> dict:
    if not os.path.exists(TEU_UPDATE_LOG_PATH):
        return {}
    with open(TEU_UPDATE_LOG_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_teu_update_log(log: dict) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(TEU_UPDATE_LOG_PATH, 'w', encoding='utf-8') as f:
        json.dump(log, f, ensure_ascii=False, indent=2)


def _find_col(columns: list, keywords: list) -> str | None:
    for col in columns:
        for kw in keywords:
            if kw in col:
                return col
    return None


def _teu_stats(sub: pd.DataFrame, teu_col: str) -> dict:
    """Calcule les stats TEU pour un sous-groupe."""
    n_containers = len(sub)
    n1 = int((sub[teu_col] == 1).sum())
    n2 = int((sub[teu_col] == 2).sum())
    teu_total = int(n1 * 1 + n2 * 2)
    return {
        'containers': n_containers,
        'teu_20':     n1,
        'teu_40':     n2,
        'teu_total':  teu_total,
    }


def _compute_teu_ratios(df: pd.DataFrame,
                         lane_col: str, dl_col: str,
                         fm_col: str, teu_col: str,
                         date_col: str) -> dict:
    """
    Calcule les ratios TEU par (lane, D/L, F/M) sur les 3 derniers mois.
    """
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=False, errors='coerce')
    df[dl_col]   = df[dl_col].astype(str).str.strip().str.upper()
    df[fm_col]   = df[fm_col].astype(str).str.strip().str.upper()
    df[teu_col]  = pd.to_numeric(df[teu_col], errors='coerce')

    df = df.dropna(subset=[date_col, lane_col, dl_col, fm_col, teu_col])
    df = df[df[dl_col].isin(['L','D']) & df[fm_col].isin(['F','M']) & df[teu_col].isin([1.0,2.0])]

    if df.empty:
        return {}

    max_date   = df[date_col].max()
    start_date = max_date - pd.DateOffset(months=MONTHS_WINDOW)
    df_w       = df[df[date_col] >= start_date].copy()

    if df_w.empty:
        return {}

    df_w[lane_col] = df_w[lane_col].astype(str).str.strip().str.upper()

    new_ratios = {}
    for lane, grp in df_w.groupby(lane_col):
        if not lane or lane in ('NAN', ''):
            continue

        lane_data = {
            'last_date':    grp[date_col].max().strftime('%Y-%m-%d'),
            'period_start': grp[date_col].min().strftime('%Y-%m-%d'),
        }

        for dl, dl_key in [('D', 'discharge'), ('L', 'load')]:
            sub_dl = grp[grp[dl_col] == dl]
            total_teu = int((sub_dl[teu_col]==1).sum()*1 + (sub_dl[teu_col]==2).sum()*2)
            dl_data = {'total_teu': total_teu}

            for fm, fm_key in [('F', 'full'), ('M', 'empty')]:
                sub_fm = sub_dl[sub_dl[fm_col] == fm]
                stats  = _teu_stats(sub_fm, teu_col)
                pct    = round(stats['teu_total'] / total_teu * 100, 2) if total_teu > 0 else 0.0
                stats['pct_teu'] = pct
                dl_data[fm_key] = stats

            lane_data[dl_key] = dl_data

        new_ratios[lane] = lane_data

    return new_ratios


def process_moves_file_teu(file_bytes: bytes, filename: str) -> dict:
    """
    Lit un fichier de moves, recalcule les ratios TEU,
    met à jour teu_ratios.json et retourne un rapport.
    """
    try:
        df = pd.read_csv(io.BytesIO(file_bytes)) if filename.lower().endswith('.csv') \
             else pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        return {'success': False, 'error': f'Impossible de lire le fichier : {e}'}

    df.columns = [str(c).strip().lower() for c in df.columns]

    lane_col = _find_col(df.columns, ['lane'])
    fm_col   = _find_col(df.columns, ['f/m'])
    teu_col  = _find_col(df.columns, ['teu'])
    date_col = _find_col(df.columns, ['dis./load/as completed date', 'completed date',
                                       'as completed', 'move date', 'date'])
    dl_candidates = [c for c in df.columns if 'd/l' in c]
    dl_col = min(dl_candidates, key=len) if dl_candidates else None

    if not lane_col: return {'success': False, 'error': "Colonne 'Lane' introuvable."}
    if not dl_col:   return {'success': False, 'error': "Colonne 'D/L' introuvable."}
    if not fm_col:   return {'success': False, 'error': "Colonne 'F/M' introuvable."}
    if not teu_col:  return {'success': False, 'error': "Colonne 'TEU' introuvable."}
    if not date_col: return {'success': False, 'error': "Colonne de date introuvable."}

    new_ratios = _compute_teu_ratios(df, lane_col, dl_col, fm_col, teu_col, date_col)
    if not new_ratios:
        return {'success': False, 'error': "Aucune donnée TEU valide trouvée."}

    existing        = load_teu_ratios()
    updated_lanes   = sorted(set(existing) & set(new_ratios))
    new_lanes       = sorted(set(new_ratios) - set(existing))
    unchanged_lanes = sorted(set(existing) - set(new_ratios))

    save_teu_ratios({**existing, **new_ratios})

    all_dates    = pd.to_datetime(df[date_col], errors='coerce').dropna()
    max_date     = all_dates.max()
    period_end   = max_date.strftime('%Y-%m-%d') if not pd.isnull(max_date) else '?'
    period_start = (max_date - pd.DateOffset(months=MONTHS_WINDOW)).strftime('%Y-%m-%d') \
                   if not pd.isnull(max_date) else '?'

    log = {
        'updated_at':       datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'filename':         filename,
        'total_rows':       len(df),
        'updated_lanes':    updated_lanes,
        'new_lanes':        new_lanes,
        'unchanged_lanes':  unchanged_lanes,
        'period_start':     period_start,
        'period_end':       period_end,
    }
    save_teu_update_log(log)

    return {
        'success': True, 'error': None,
        'updated_lanes': updated_lanes, 'new_lanes': new_lanes,
        'unchanged_lanes': unchanged_lanes,
        'total_rows': len(df),
        'period_start': period_start, 'period_end': period_end,
    }
