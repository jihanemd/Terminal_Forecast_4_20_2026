from flask import Flask, render_template, request, jsonify, send_file
from utils.predictor import predict_single, predict_batch, reload_ratios
from utils.export import export_to_excel, export_fm_to_excel
from utils.updater import process_moves_file, load_update_log
from utils.fm_predictor import predict_full_pipeline, predict_fm_batch, reload_fm_ratios
from utils.updater_fm import process_moves_file_fm, load_fm_update_log
from utils.teu_predictor import predict_full_pipeline_with_teu, predict_teu_batch, reload_teu_ratios
from utils.updater_teu import process_moves_file_teu, load_teu_update_log
from datetime import datetime
import os, io

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB max upload

# ── Pages ──────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')

# ── API ────────────────────────────────────────────────────────────────────────

@app.route('/api/predict', methods=['POST'])
def api_predict():
    """Prédiction pour une seule lane."""
    data = request.get_json()
    lane   = str(data.get('lane', '')).strip().upper()
    volume = int(data.get('volume', 0))

    if not lane:
        return jsonify({'error': 'Lane manquante'}), 400
    if volume <= 0:
        return jsonify({'error': 'Volume invalide'}), 400

    result = predict_single(lane, volume)
    if result is None:
        return jsonify({'error': f'Lane "{lane}" inconnue'}), 404

    return jsonify(result)


@app.route('/api/predict/batch', methods=['POST'])
def api_predict_batch():
    """Prédiction pour un fichier Excel ou CSV uploadé."""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier reçu'}), 400

    file = request.files['file']
    filename = file.filename.lower()

    if not (filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.csv')):
        return jsonify({'error': 'Format non supporté. Utilisez .xlsx, .xls ou .csv'}), 400

    file_bytes = file.read()
    results, errors = predict_batch(file_bytes, filename)

    if errors and not results:
        return jsonify({'error': errors[0]}), 400

    return jsonify({'results': results, 'errors': errors, 'count': len(results)})


@app.route('/api/export', methods=['POST'])
def api_export():
    """Exporte les résultats batch en fichier Excel."""
    data = request.get_json()
    results = data.get('results', [])

    if not results:
        return jsonify({'error': 'Aucun résultat à exporter'}), 400

    excel_buffer = export_to_excel(results)
    return send_file(
        excel_buffer,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='predictions_LD.xlsx'
    )


@app.route('/api/lanes', methods=['GET'])
def api_lanes():
    """Retourne la liste de toutes les lanes disponibles."""
    from utils.predictor import RATIOS
    lanes = sorted(RATIOS.keys())
    return jsonify({'lanes': lanes, 'count': len(lanes)})


@app.route('/api/update-ratios', methods=['POST'])
def api_update_ratios():
    """Upload d'un fichier de moves mensuel pour recalculer et mettre à jour les RATIOS."""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier reçu'}), 400

    file     = request.files['file']
    filename = file.filename

    if not filename:
        return jsonify({'error': 'Nom de fichier manquant'}), 400

    ext = filename.lower()
    if not (ext.endswith('.xlsx') or ext.endswith('.xls') or ext.endswith('.csv')):
        return jsonify({'error': 'Format non supporté. Utilisez .xlsx, .xls ou .csv'}), 400

    file_bytes = file.read()
    result = process_moves_file(file_bytes, filename)

    if not result['success']:
        return jsonify({'error': result['error']}), 400

    # Recharger les RATIOS en mémoire immédiatement
    n_lanes = reload_ratios()

    return jsonify({
        'success':         True,
        'message':         f'RATIOS mis à jour avec succès. {n_lanes} lanes actives.',
        'updated_lanes':   result['updated_lanes'],
        'new_lanes':       result['new_lanes'],
        'unchanged_lanes': result['unchanged_lanes'],
        'total_rows':      result['total_rows'],
        'period_start':    result['period_start'],
        'period_end':      result['period_end'],
        'n_lanes':         n_lanes,
    })


@app.route('/api/update-status', methods=['GET'])
def api_update_status():
    """Retourne les informations de la dernière mise à jour des RATIOS."""
    log = load_update_log()
    from utils.predictor import RATIOS
    return jsonify({
        'last_update':    log.get('updated_at', None),
        'filename':       log.get('filename', None),
        'n_lanes':        len(RATIOS),
        'period_start':   log.get('period_start', None),
        'period_end':     log.get('period_end', None),
        'updated_lanes':  log.get('updated_lanes', []),
        'new_lanes':      log.get('new_lanes', []),
    })


@app.route('/api/dashboard', methods=['GET'])
def api_dashboard():
    """Retourne les KPIs globaux du système."""
    from utils.predictor import RATIOS
    now = datetime.now()
    fresh = stale = outdated = 0
    most_loaded_lane = None
    max_total = 0
    for lane, r in RATIOS.items():
        try:
            months = (now - datetime.strptime(r['last_date'], '%Y-%m-%d')).days / 30
            if months < 6:   fresh    += 1
            elif months < 12: stale   += 1
            else:             outdated += 1
        except:
            outdated += 1
        if r.get('total', 0) > max_total:
            max_total = r.get('total', 0)
            most_loaded_lane = lane
    log = load_update_log()
    return jsonify({
        'n_lanes':          len(RATIOS),
        'fresh_lanes':      fresh,
        'stale_lanes':      stale,
        'outdated_lanes':   outdated,
        'most_loaded_lane': most_loaded_lane,
        'last_update':      log.get('updated_at', None),
    })


@app.route('/api/lanes/details', methods=['GET'])
def api_lanes_details():
    """Retourne toutes les lanes avec indicateur de fraîcheur."""
    from utils.predictor import RATIOS
    now = datetime.now()
    lanes = []
    for lane, r in sorted(RATIOS.items()):
        try:
            months = (now - datetime.strptime(r['last_date'], '%Y-%m-%d')).days / 30
            freshness = 'fresh' if months < 6 else ('stale' if months < 12 else 'outdated')
        except:
            freshness = 'outdated'
        lanes.append({'lane': lane, 'freshness': freshness, **r})
    return jsonify({'lanes': lanes})


@app.route('/admin')
def admin():
    """Page d'administration des lanes."""
    return render_template('admin.html')


@app.route('/api/predict/compare', methods=['POST'])
def api_predict_compare():
    """Prédiction comparée pour plusieurs lanes avec le même volume."""
    data   = request.get_json()
    lanes  = data.get('lanes', [])
    volume = int(data.get('volume', 0))
    if not lanes or volume <= 0:
        return jsonify({'error': 'Données invalides'}), 400
    results = []
    for lane in lanes:
        lane = str(lane).strip().upper()
        res  = predict_single(lane, volume)
        results.append(res if res else {'lane': lane, 'unknown': True})
    return jsonify({'results': results, 'volume': volume})


# ── Full / Empty routes ────────────────────────────────────────────────────────

@app.route('/api/predict/fm', methods=['POST'])
def api_predict_fm():
    """Pipeline complet : Lane + Volume → D/L puis F/M."""
    data   = request.get_json()
    lane   = str(data.get('lane', '')).strip().upper()
    volume = int(data.get('volume', 0))
    if not lane:   return jsonify({'error': 'Lane manquante'}), 400
    if volume <= 0: return jsonify({'error': 'Volume invalide'}), 400
    result = predict_full_pipeline(lane, volume)
    if result is None:
        return jsonify({'error': f'Lane "{lane}" inconnue'}), 404
    return jsonify(result)


@app.route('/api/predict/fm/batch', methods=['POST'])
def api_predict_fm_batch():
    """Pipeline complet batch depuis un fichier Excel/CSV."""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier reçu'}), 400
    file     = request.files['file']
    filename = file.filename.lower()
    if not (filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.csv')):
        return jsonify({'error': 'Format non supporté. Utilisez .xlsx, .xls ou .csv'}), 400
    file_bytes      = file.read()
    results, errors = predict_fm_batch(file_bytes, filename)
    if errors and not results:
        return jsonify({'error': errors[0]}), 400
    return jsonify({'results': results, 'errors': errors, 'count': len(results)})


@app.route('/api/export/fm', methods=['POST'])
def api_export_fm():
    """Exporte les résultats F/M batch en fichier Excel."""
    data    = request.get_json()
    results = data.get('results', [])
    if not results:
        return jsonify({'error': 'Aucun résultat à exporter'}), 400
    excel_buffer = export_fm_to_excel(results)
    return send_file(excel_buffer,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='predictions_FM.xlsx')


@app.route('/api/update-fm-ratios', methods=['POST'])
def api_update_fm_ratios():
    """Upload d'un fichier de moves pour recalculer les ratios F/M."""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier reçu'}), 400
    file     = request.files['file']
    filename = file.filename
    if not filename:
        return jsonify({'error': 'Nom de fichier manquant'}), 400
    ext = filename.lower()
    if not (ext.endswith('.xlsx') or ext.endswith('.xls') or ext.endswith('.csv')):
        return jsonify({'error': 'Format non supporté. Utilisez .xlsx, .xls ou .csv'}), 400
    result  = process_moves_file_fm(file.read(), filename)
    if not result['success']:
        return jsonify({'error': result['error']}), 400
    n_lanes = reload_fm_ratios()
    return jsonify({'success': True,
                    'message': f'Ratios F/M mis à jour. {n_lanes} lanes actives.',
                    'updated_lanes': result['updated_lanes'],
                    'new_lanes':     result['new_lanes'],
                    'unchanged_lanes': result['unchanged_lanes'],
                    'total_rows':    result['total_rows'],
                    'period_start':  result['period_start'],
                    'period_end':    result['period_end'],
                    'n_lanes':       n_lanes})


@app.route('/api/update-fm-status', methods=['GET'])
def api_update_fm_status():
    """Statut de la dernière mise à jour des ratios F/M."""
    log = load_fm_update_log()
    from utils.fm_predictor import FM_RATIOS
    return jsonify({'last_update':  log.get('updated_at', None),
                    'filename':     log.get('filename', None),
                    'n_lanes':      len(FM_RATIOS),
                    'period_start': log.get('period_start', None),
                    'period_end':   log.get('period_end', None),
                    'updated_lanes': log.get('updated_lanes', []),
                    'new_lanes':    log.get('new_lanes', [])})


# ── Run ────────────────────────────────────────────────────────────────────────

# ── TEU routes ─────────────────────────────────────────────────────────────────

@app.route('/api/predict/teu', methods=['POST'])
def api_predict_teu():
    """Pipeline complet 3 étapes : Lane + Volume → D/L → F/M → TEU."""
    data   = request.get_json()
    lane   = str(data.get('lane', '')).strip().upper()
    volume = int(data.get('volume', 0))
    if not lane:    return jsonify({'error': 'Lane manquante'}), 400
    if volume <= 0: return jsonify({'error': 'Volume invalide'}), 400
    result = predict_full_pipeline_with_teu(lane, volume)
    if result is None:
        return jsonify({'error': f'Lane "{lane}" inconnue'}), 404
    return jsonify(result)


@app.route('/api/predict/teu/batch', methods=['POST'])
def api_predict_teu_batch():
    """Pipeline complet 3 étapes batch depuis un fichier Excel/CSV."""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier reçu'}), 400
    file     = request.files['file']
    filename = file.filename.lower()
    if not (filename.endswith('.xlsx') or filename.endswith('.xls') or filename.endswith('.csv')):
        return jsonify({'error': 'Format non supporté. Utilisez .xlsx, .xls ou .csv'}), 400
    results, errors = predict_teu_batch(file.read(), filename)
    if errors and not results:
        return jsonify({'error': errors[0]}), 400
    return jsonify({'results': results, 'errors': errors, 'count': len(results)})


@app.route('/api/export/teu', methods=['POST'])
def api_export_teu():
    """Exporte les résultats TEU complets en fichier Excel."""
    data    = request.get_json()
    results = data.get('results', [])
    if not results:
        return jsonify({'error': 'Aucun résultat à exporter'}), 400
    from utils.export import export_teu_to_excel
    excel_buffer = export_teu_to_excel(results)
    return send_file(excel_buffer,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='predictions_TEU.xlsx')


@app.route('/api/update-teu-ratios', methods=['POST'])
def api_update_teu_ratios():
    """Upload d'un fichier de moves pour recalculer les ratios TEU."""
    if 'file' not in request.files:
        return jsonify({'error': 'Aucun fichier reçu'}), 400
    file     = request.files['file']
    filename = file.filename
    if not filename:
        return jsonify({'error': 'Nom de fichier manquant'}), 400
    ext = filename.lower()
    if not (ext.endswith('.xlsx') or ext.endswith('.xls') or ext.endswith('.csv')):
        return jsonify({'error': 'Format non supporté.'}), 400
    result  = process_moves_file_teu(file.read(), filename)
    if not result['success']:
        return jsonify({'error': result['error']}), 400
    n_lanes = reload_teu_ratios()
    return jsonify({'success': True,
                    'message': f'Ratios TEU mis à jour. {n_lanes} lanes actives.',
                    'updated_lanes': result['updated_lanes'],
                    'new_lanes':     result['new_lanes'],
                    'unchanged_lanes': result['unchanged_lanes'],
                    'total_rows':    result['total_rows'],
                    'period_start':  result['period_start'],
                    'period_end':    result['period_end'],
                    'n_lanes':       n_lanes})


@app.route('/api/update-teu-status', methods=['GET'])
def api_update_teu_status():
    """Statut de la dernière mise à jour des ratios TEU."""
    log = load_teu_update_log()
    from utils.teu_predictor import TEU_RATIOS
    return jsonify({'last_update':  log.get('updated_at', None),
                    'filename':     log.get('filename', None),
                    'n_lanes':      len(TEU_RATIOS),
                    'period_start': log.get('period_start', None),
                    'period_end':   log.get('period_end', None),
                    'updated_lanes': log.get('updated_lanes', []),
                    'new_lanes':    log.get('new_lanes', [])})


# ── Run ─────────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
