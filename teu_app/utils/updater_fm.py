"""
updater_fm.py
-------------
Mise à jour mensuelle des ratios F/M depuis un fichier de moves.
Calcule les ratios Full/Empty séparément pour Discharge et Load par lane,
sur les 3 derniers mois avant la date max du fichier.

Colonnes attendues : Lane · D/L (ou D/L/S) · F/M · date
Fichier de sortie  : data/fm_ratios.json
Log                : data/last_update_fm.json
"""

import os, json, io
from datetime import datetime
import pandas as pd

DATA_DIR           = os.path.join(os.path.dirname(__file__), '..', 'data')
FM_RATIOS_PATH     = os.path.join(DATA_DIR, 'fm_ratios.json')
FM_UPDATE_LOG_PATH = os.path.join(DATA_DIR, 'last_update_fm.json')
MONTHS_WINDOW      = 3


def load_fm_ratios() -> dict:
    if not os.path.exists(FM_RATIOS_PATH):
        return {}
    with open(FM_RATIOS_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_fm_ratios(ratios: dict) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(FM_RATIOS_PATH, 'w', encoding='utf-8') as f:
        json.dump(ratios, f, ensure_ascii=False, indent=2)

def load_fm_update_log() -> dict:
    if not os.path.exists(FM_UPDATE_LOG_PATH):
        return {}
    with open(FM_UPDATE_LOG_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_fm_update_log(log: dict) -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(FM_UPDATE_LOG_PATH, 'w', encoding='utf-8') as f:
        json.dump(log, f, ensure_ascii=False, indent=2)


def _find_col(columns: list, keywords: list) -> str | None:
    for col in columns:
        for kw in keywords:
            if kw in col:
                return col
    return None


def _compute_fm_ratios(df, lane_col, dl_col, fm_col, date_col) -> dict:
    df[date_col] = pd.to_datetime(df[date_col], dayfirst=False, errors='coerce')
    df[dl_col]   = df[dl_col].astype(str).str.strip().str.upper()
    df[fm_col]   = df[fm_col].astype(str).str.strip().str.upper()
    df = df.dropna(subset=[date_col, lane_col, dl_col, fm_col])
    df = df[df[dl_col].isin(['L', 'D']) & df[fm_col].isin(['F', 'M'])]
    if df.empty:
        return {}

    max_date   = df[date_col].max()
    start_date = max_date - pd.DateOffset(months=MONTHS_WINDOW)
    df_w       = df[df[date_col] >= start_date].copy()
    if df_w.empty:
        return {}

    df_w[lane_col] = df_w[lane_col].astype(str).str.strip().str.upper()

    def _stats(sub):
        total_F = int((sub[fm_col] == 'F').sum())
        total_M = int((sub[fm_col] == 'M').sum())
        total   = total_F + total_M
        if total == 0:
            return None
        return {'total_F': total_F, 'total_M': total_M, 'total': total,
                'pct_F': round(total_F / total * 100, 2),
                'pct_M': round(total_M / total * 100, 2)}

    new_ratios = {}
    for lane, grp in df_w.groupby(lane_col):
        if not lane or lane in ('NAN', ''):
            continue
        rd = _stats(grp[grp[dl_col] == 'D'])
        rl = _stats(grp[grp[dl_col] == 'L'])
        if rd is None and rl is None:
            continue
        new_ratios[lane] = {
            'last_date':    grp[date_col].max().strftime('%Y-%m-%d'),
            'period_start': grp[date_col].min().strftime('%Y-%m-%d'),
            'discharge':    rd,
            'load':         rl,
        }
    return new_ratios


def process_moves_file_fm(file_bytes: bytes, filename: str) -> dict:
    try:
        df = pd.read_csv(io.BytesIO(file_bytes)) if filename.lower().endswith('.csv') \
             else pd.read_excel(io.BytesIO(file_bytes))
    except Exception as e:
        return {'success': False, 'error': f'Impossible de lire le fichier : {e}'}

    df.columns = [str(c).strip().lower() for c in df.columns]

    lane_col = _find_col(df.columns, ['lane'])
    fm_col   = _find_col(df.columns, ['f/m'])
    date_col = _find_col(df.columns, ['dis./load/as completed date', 'completed date',
                                       'as completed', 'move date', 'date'])
    dl_candidates = [c for c in df.columns if 'd/l' in c]
    dl_col = min(dl_candidates, key=len) if dl_candidates else None

    if not lane_col: return {'success': False, 'error': "Colonne 'Lane' introuvable."}
    if not dl_col:   return {'success': False, 'error': "Colonne 'D/L' ou 'D/L/S' introuvable."}
    if not fm_col:   return {'success': False, 'error': "Colonne 'F/M' introuvable."}
    if not date_col: return {'success': False, 'error': "Colonne de date introuvable."}

    new_ratios = _compute_fm_ratios(df, lane_col, dl_col, fm_col, date_col)
    if not new_ratios:
        return {'success': False, 'error': "Aucune donnée F/M valide trouvée."}

    existing        = load_fm_ratios()
    updated_lanes   = sorted(set(existing) & set(new_ratios))
    new_lanes       = sorted(set(new_ratios) - set(existing))
    unchanged_lanes = sorted(set(existing) - set(new_ratios))

    save_fm_ratios({**existing, **new_ratios})

    all_dates    = pd.to_datetime(df[date_col], errors='coerce').dropna()
    max_date     = all_dates.max()
    period_end   = max_date.strftime('%Y-%m-%d') if not pd.isnull(max_date) else '?'
    period_start = (max_date - pd.DateOffset(months=MONTHS_WINDOW)).strftime('%Y-%m-%d') \
                   if not pd.isnull(max_date) else '?'

    log = {'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
           'filename': filename, 'total_rows': len(df),
           'updated_lanes': updated_lanes, 'new_lanes': new_lanes,
           'unchanged_lanes': unchanged_lanes,
           'period_start': period_start, 'period_end': period_end}
    save_fm_update_log(log)

    return {'success': True, 'error': None,
            'updated_lanes': updated_lanes, 'new_lanes': new_lanes,
            'unchanged_lanes': unchanged_lanes, 'total_rows': len(df),
            'period_start': period_start, 'period_end': period_end}
