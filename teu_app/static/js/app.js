'use strict';

// ── State ────────────────────────────────────────────────────────────────────
let allLanes       = [];   // [{lane, freshness, last_date, ...}]
let selectedLane   = '';
let sessionHistory = [];
let compareLanes   = [];
let lastBatchResults = null;

// ── Init ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  initDarkMode();
  loadLanesDetails();
  loadDashboard();
  loadUpdateStatus();
  initComparePanel();
});

// ── Dark mode ────────────────────────────────────────────────────────────────
function initDarkMode() {
  const stored = localStorage.getItem('theme') || 'light';
  applyTheme(stored);
}
function toggleDarkMode() {
  const current = document.documentElement.getAttribute('data-theme');
  applyTheme(current === 'dark' ? 'light' : 'dark');
}
function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('theme', theme);
  document.getElementById('iconMoon').style.display = theme === 'dark'  ? 'none' : '';
  document.getElementById('iconSun').style.display  = theme === 'light' ? 'none' : '';
}

// ── Toast ────────────────────────────────────────────────────────────────────
function showToast(msg, type = 'info') {
  const c   = document.getElementById('toastContainer');
  const div = document.createElement('div');
  div.className = `toast toast-${type}`;
  div.textContent = msg;
  c.appendChild(div);
  setTimeout(() => { div.classList.add('hide'); setTimeout(() => div.remove(), 260); }, 3500);
}

// ── Dashboard KPIs ───────────────────────────────────────────────────────────
async function loadDashboard() {
  try {
    const d = await fetch('/api/dashboard').then(r => r.json());
    animateNum('kpiLanes',    d.n_lanes);
    animateNum('kpiFresh',    d.fresh_lanes);
    animateNum('kpiStale',    d.stale_lanes);
    animateNum('kpiOutdated', d.outdated_lanes);
  } catch(e) { console.error('Dashboard KPI error', e); }
}
function animateNum(id, target) {
  const el = document.getElementById(id);
  if (!el) return;
  let start = 0;
  const step = Math.ceil(target / 20);
  const iv = setInterval(() => {
    start = Math.min(start + step, target);
    el.textContent = start;
    if (start >= target) clearInterval(iv);
  }, 40);
}

// ── Lanes (with details) ─────────────────────────────────────────────────────
async function loadLanesDetails() {
  try {
    const d = await fetch('/api/lanes/details').then(r => r.json());
    allLanes = d.lanes;
    buildComboList(allLanes);
    // Sidebar + footer count
    const c1 = document.getElementById('lanesCount');
    const c2 = document.getElementById('sidebarLaneCount');
    if (c1) c1.textContent = allLanes.length;
    if (c2) c2.textContent = allLanes.length;
  } catch(e) { console.error('Lanes load error', e); }
}

// ── Combobox ─────────────────────────────────────────────────────────────────
function buildComboList(lanes) {
  const list = document.getElementById('comboList');
  list.innerHTML = '';
  if (!lanes.length) {
    list.innerHTML = '<div class="combo-empty">Aucune lane trouvée</div>';
    return;
  }
  lanes.forEach(l => {
    const div = document.createElement('div');
    div.className = 'combo-option';
    div.dataset.lane = l.lane;
    div.innerHTML = `
      <span class="opt-dot" style="background:${freshColor(l.freshness)}"></span>
      <span class="opt-name">${l.lane}</span>
      <span class="opt-date">${l.last_date}</span>`;
    div.addEventListener('mousedown', (e) => { e.preventDefault(); selectLane(l); });
    list.appendChild(div);
  });
}
function freshColor(f) {
  return f === 'fresh' ? '#10B981' : f === 'stale' ? '#F59E0B' : '#EF4444';
}
function toggleCombo() {
  const cb = document.getElementById('laneCombobox');
  const dd = document.getElementById('comboDropdown');
  const open = cb.classList.toggle('open');
  dd.style.display = open ? 'block' : 'none';
  if (open) { document.getElementById('comboSearch').focus(); filterCombo(); }
}
function filterCombo() {
  const q = document.getElementById('comboSearch').value.toLowerCase();
  const filtered = allLanes.filter(l => l.lane.toLowerCase().includes(q));
  buildComboList(filtered);
}
function selectLane(l) {
  selectedLane = l.lane;
  document.getElementById('selectedLane').value    = l.lane;
  document.getElementById('comboText').textContent = l.lane;
  document.getElementById('comboDot').className    = `combo-dot ${l.freshness}`;
  document.getElementById('comboDropdown').style.display = 'none';
  document.getElementById('laneCombobox').classList.remove('open');
  document.getElementById('comboSearch').value = '';
  document.getElementById('laneText').value    = '';
}
// Close combo on outside click
document.addEventListener('click', (e) => {
  if (!document.getElementById('laneCombobox')?.contains(e.target)) {
    document.getElementById('comboDropdown').style.display = 'none';
    document.getElementById('laneCombobox')?.classList.remove('open');
  }
});

// ── Page switch (sidebar) ───────────────────────────────────────────────────
function setPage(page) {
  // Hide all panels
  document.querySelectorAll('.page-panel').forEach(p => p.classList.remove('active'));
  // Show target
  const target = document.getElementById('page-' + page);
  if (target) target.classList.add('active');
  // Update nav items
  document.querySelectorAll('.nav-item').forEach(b => {
    b.classList.toggle('active', b.id === 'nav-' + page);
  });
}
// Legacy alias for any remaining calls
function setMode(m) { setPage(m === 'manual' ? 'predict' : m === 'excel' ? 'import' : m); }

// ── Lane text sync ───────────────────────────────────────────────────────────
function onLaneText() {
  const v = document.getElementById('laneText').value.toUpperCase();
  document.getElementById('laneText').value = v;
  if (v) {
    selectedLane = v;
    document.getElementById('comboText').textContent = '— Choisir —';
    document.getElementById('comboDot').className = 'combo-dot';
    document.getElementById('selectedLane').value = '';
  }
}
function getManualLane() {
  return document.getElementById('selectedLane').value || document.getElementById('laneText').value.trim().toUpperCase() || null;
}

// ── Helpers ──────────────────────────────────────────────────────────────────
const fmt = n => n != null ? Math.round(n).toLocaleString('fr-FR') : '—';

function setLoading(id, msg = 'Calcul en cours…') {
  document.getElementById(id).innerHTML =
    `<div style="padding:16px 0;font-size:13px;color:var(--muted)"><span class="spinner dark"></span>${msg}</div>`;
}

// ── Manual predict ───────────────────────────────────────────────────────────
async function predict() {
  const lane   = getManualLane();
  const volume = parseInt(document.getElementById('volumeInput').value) || 0;
  if (!lane)   { showToast('Veuillez sélectionner ou saisir une lane.', 'warn'); return; }
  if (!volume) { showToast('Veuillez saisir un volume valide.', 'warn'); return; }

  const btn = document.getElementById('btnPredict');
  btn.classList.add('loading');
  btn.textContent = 'Calcul…';
  setLoading('resultManual');

  try {
    const res  = await fetch('/api/predict', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({lane,volume}) });
    const data = await res.json();
    btn.classList.remove('loading'); btn.textContent = 'Calculer Load & Discharge';
    if (!res.ok) {
      document.getElementById('resultManual').innerHTML = `<div class="error-box">❌ ${data.error}</div>`;
      showToast(data.error, 'error');
      return;
    }
    document.getElementById('resultManual').innerHTML = buildResultHTML(data);
    addToHistory(data);
    showToast(`Lane ${data.lane} — L: ${fmt(data.pred_L)} · D: ${fmt(data.pred_D)}`, 'success');
  } catch(e) {
    btn.classList.remove('loading'); btn.textContent = 'Calculer Load & Discharge';
    document.getElementById('resultManual').innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`;
    showToast('Erreur réseau', 'error');
  }
}

function buildResultHTML(d) {
  const circ = 2 * Math.PI * 30;
  const loadArc = (d.pct_L / 100) * circ;
  const dischArc = (d.pct_D / 100) * circ;
  const warn = d.outdated
    ? `<div class="warn-box">⚠ Dernière escale le <strong>${d.last_date}</strong> — données potentiellement obsolètes</div>` : '';

  return `
    <div class="result-section">
      <div class="result-header">Lane ${d.lane} &nbsp;·&nbsp; Volume ${fmt(d.volume)}</div>
      ${warn}
      <div class="result-grid">
        <div class="result-card load">
          <div class="rc-label">Load estimé</div>
          <div class="rc-value" id="rcL">${fmt(d.pred_L)}</div>
          <div class="rc-pct">${d.pct_L.toFixed(1)}% du volume</div>
        </div>
        <div class="result-card disch">
          <div class="rc-label">Discharge estimé</div>
          <div class="rc-value" id="rcD">${fmt(d.pred_D)}</div>
          <div class="rc-pct">${d.pct_D.toFixed(1)}% du volume</div>
        </div>
      </div>
      <div class="donut-wrap">
        <svg class="donut-svg" viewBox="0 0 80 80">
          <circle class="donut-track" cx="40" cy="40" r="30"/>
          <circle class="donut-load" cx="40" cy="40" r="30"
            stroke-dasharray="${loadArc} ${circ}"
            stroke-dashoffset="0"
            transform="rotate(-90 40 40)"/>
          <circle class="donut-disch" cx="40" cy="40" r="30"
            stroke-dasharray="${dischArc} ${circ}"
            stroke-dashoffset="${-loadArc}"
            transform="rotate(-90 40 40)"/>
          <text class="donut-text" x="40" y="37">L / D</text>
          <text class="donut-text" x="40" y="50" style="font-size:10px;fill:var(--ink);font-weight:500">${d.pct_L.toFixed(0)}/${d.pct_D.toFixed(0)}</text>
        </svg>
        <div class="donut-legend">
          <div class="donut-legend-item"><span class="dl-dot" style="background:var(--blue)"></span>Load<span class="dl-val" style="color:var(--blue)">${d.pct_L.toFixed(1)}%</span></div>
          <div class="donut-legend-item"><span class="dl-dot" style="background:var(--coral)"></span>Discharge<span class="dl-val" style="color:var(--coral)">${d.pct_D.toFixed(1)}%</span></div>
          <div class="donut-legend-item" style="margin-top:4px;border-top:1px solid var(--border);padding-top:8px">
            <span style="font-size:11px;color:var(--muted);font-family:var(--mono)">${d.period_start} → ${d.last_date}</span>
          </div>
        </div>
      </div>
      <div class="info-grid">
        <div class="info-cell"><div class="ic-label">Escales analysées</div><div class="ic-value">${d.n_voyages} voyages</div></div>
        <div class="info-cell"><div class="ic-label">Moy. Load / escale</div><div class="ic-value">${fmt(d.avg_L)}</div></div>
        <div class="info-cell"><div class="ic-label">Moy. Discharge / escale</div><div class="ic-value">${fmt(d.avg_D)}</div></div>
        <div class="info-cell"><div class="ic-label">Total analysé</div><div class="ic-value">${fmt(d.pred_L + d.pred_D)}</div></div>
      </div>
    </div>`;
}

// ── Session History ──────────────────────────────────────────────────────────
function addToHistory(result) {
  sessionHistory.unshift(result);
  if (sessionHistory.length > 10) sessionHistory.pop();
  renderHistory();
}
function renderHistory() {
  const card = document.getElementById('historyCard');
  const list = document.getElementById('historyList');
  if (!sessionHistory.length) { card.style.display = 'none'; return; }
  card.style.display = 'block';
  list.innerHTML = sessionHistory.map((r, i) => {
    const laneData = allLanes.find(l => l.lane === r.lane);
    const dot = laneData ? freshColor(laneData.freshness) : 'var(--border)';
    return `
      <div class="history-item" onclick="replayHistory(${i})">
        <span class="hi-dot" style="background:${dot}"></span>
        <span class="hi-lane">${r.lane}</span>
        <span class="hi-vol">${fmt(r.volume)} mvt</span>
        <div class="hi-ld">
          <span class="hi-l">L ${fmt(r.pred_L)}</span>
          <span class="hi-d">D ${fmt(r.pred_D)}</span>
        </div>
      </div>`;
  }).join('');
}
function replayHistory(i) {
  const r = sessionHistory[i];
  setMode('manual');
  document.getElementById('laneText').value    = r.lane;
  document.getElementById('volumeInput').value = r.volume;
  selectedLane = r.lane;
  document.getElementById('resultManual').innerHTML = buildResultHTML(r);
}
function clearHistory() {
  sessionHistory = [];
  document.getElementById('historyCard').style.display = 'none';
}

// ── Compare mode ─────────────────────────────────────────────────────────────
function initComparePanel() {
  addCompareLane(); addCompareLane();
}
function addCompareLane() {
  if (compareLanes.length >= 4) { showToast('Maximum 4 lanes pour la comparaison.', 'warn'); return; }
  compareLanes.push('');
  renderCompareLanes();
}
function removeCompareLane(i) {
  compareLanes.splice(i, 1);
  renderCompareLanes();
}
function renderCompareLanes() {
  const row = document.getElementById('compareLanesRow');
  const addBtn = document.getElementById('btnAddLane');
  row.innerHTML = compareLanes.map((l, i) => {
    const laneData = allLanes.find(x => x.lane === l);
    const dot = laneData ? freshColor(laneData.freshness) : 'transparent';
    const border = laneData ? `2px solid ${freshColor(laneData.freshness)}` : '1px dashed var(--border)';
    return `
      <div class="compare-lane-item" style="border:${border}">
        <span class="clane-dot" style="background:${dot}"></span>
        <select class="clane-select" style="border:none;background:transparent;font-family:var(--mono);font-size:13px;font-weight:500;outline:none;color:var(--ink)" onchange="updateCompareLane(${i}, this.value)">
          <option value="">— Lane —</option>
          ${allLanes.map(x => `<option value="${x.lane}" ${x.lane===l?'selected':''}>${x.lane}</option>`).join('')}
        </select>
        <button class="clane-rm" onclick="removeCompareLane(${i})">×</button>
      </div>`;
  }).join('');
  addBtn.style.display = compareLanes.length >= 4 ? 'none' : '';
}
function updateCompareLane(i, val) {
  compareLanes[i] = val;
  renderCompareLanes();
}
async function runComparison() {
  const lanes  = compareLanes.filter(l => l);
  const volume = parseInt(document.getElementById('compareVolume').value) || 0;
  if (lanes.length < 2) { showToast('Sélectionnez au moins 2 lanes.', 'warn'); return; }
  if (!volume)          { showToast('Saisissez un volume valide.', 'warn'); return; }
  setLoading('compareResult', 'Comparaison en cours…');
  try {
    const res  = await fetch('/api/predict/compare', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({lanes,volume}) });
    const data = await res.json();
    if (!res.ok) { document.getElementById('compareResult').innerHTML = `<div class="error-box">❌ ${data.error}</div>`; return; }
    renderComparison(data);
  } catch(e) {
    document.getElementById('compareResult').innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`;
  }
}
function renderComparison(data) {
  const rows = data.results.map(r => {
    if (r.unknown) return `<tr><td class="td-lane">${r.lane}</td><td colspan="5"><span class="pill pill-unknown">Lane inconnue</span></td></tr>`;
    return `
      <tr>
        <td class="td-lane">${r.lane}</td>
        <td><span class="pill pill-l">${fmt(r.pred_L)}</span></td>
        <td><span class="pill pill-d">${fmt(r.pred_D)}</span></td>
        <td style="font-family:var(--mono);color:var(--blue)">${r.pct_L.toFixed(1)}%</td>
        <td style="font-family:var(--mono);color:var(--coral)">${r.pct_D.toFixed(1)}%</td>
        <td class="compare-bar-cell">
          <div class="compare-bar">
            <div class="cb-l" style="width:${r.pct_L}%"></div>
            <div class="cb-d" style="width:${r.pct_D}%"></div>
          </div>
        </td>
      </tr>`;
  }).join('');
  document.getElementById('compareResult').innerHTML = `
    <div class="compare-table-wrap">
      <table>
        <thead><tr><th>Lane</th><th>Load</th><th>Discharge</th><th>% L</th><th>% D</th><th>Ratio</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

// ── File import (batch) ──────────────────────────────────────────────────────
function handleDrop(event) {
  event.preventDefault();
  document.getElementById('dropZone').classList.remove('drag-over');
  const file = event.dataTransfer.files[0];
  if (file) uploadFile(file);
}
function handleFile(event) { const f = event.target.files[0]; if (f) uploadFile(f); event.target.value=''; }
async function uploadFile(file) {
  const name = file.name.toLowerCase();
  if (!name.endsWith('.xlsx') && !name.endsWith('.xls') && !name.endsWith('.csv')) {
    document.getElementById('batchResult').innerHTML = '<div class="warn-box">Format non supporté.</div>'; return;
  }
  setLoading('batchResult', `Traitement de "${file.name}"…`);
  const fd = new FormData(); fd.append('file', file);
  try {
    const res  = await fetch('/api/predict/batch', { method:'POST', body:fd });
    const data = await res.json();
    if (!res.ok) { document.getElementById('batchResult').innerHTML = `<div class="warn-box">${data.error}</div>`; return; }
    renderBatch(data);
    showToast(`${data.count} lignes traitées avec succès`, 'success');
  } catch(e) {
    document.getElementById('batchResult').innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`;
  }
}
function renderBatch(data) {
  lastBatchResults = data.results;
  const { results, errors, count } = data;
  const errHtml = errors.length ? `<div class="warn-box" style="margin-bottom:12px">${errors.length} ligne(s) ignorée(s) : ${errors.slice(0,3).join(' | ')}</div>` : '';
  const rows = results.map(r => r.unknown
    ? `<tr><td class="td-lane">${r.lane}</td><td>${fmt(r.volume)}</td><td colspan="5"><span class="pill pill-unknown">Lane inconnue</span></td></tr>`
    : `<tr>
        <td class="td-lane">${r.lane}${r.outdated?' ⚠':''}</td>
        <td style="font-family:var(--mono)">${fmt(r.volume)}</td>
        <td><span class="pill pill-l">${fmt(r.pred_L)}</span></td>
        <td><span class="pill pill-d">${fmt(r.pred_D)}</span></td>
        <td style="font-family:var(--mono);color:var(--blue)">${r.pct_L.toFixed(1)}%</td>
        <td style="font-family:var(--mono);color:var(--coral)">${r.pct_D.toFixed(1)}%</td>
        <td style="font-family:var(--mono);font-size:11px;color:var(--muted)">${r.period_start} → ${r.last_date}</td>
       </tr>`).join('');
  document.getElementById('batchResult').innerHTML = `
    ${errHtml}
    <div class="batch-meta">
      <span>${count} ligne(s) traitée(s)</span>
      <div style="display:flex;gap:8px">
        <button class="btn-secondary" onclick="exportResults()">Exporter .xlsx</button>
        <button class="btn-secondary" onclick="window.print()">Exporter PDF</button>
      </div>
    </div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>Lane</th><th>Volume</th><th>Load</th><th>Discharge</th><th>% L</th><th>% D</th><th>Période</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

// ── Export Excel ─────────────────────────────────────────────────────────────
async function exportResults() {
  if (!lastBatchResults) return;
  try {
    const res   = await fetch('/api/export', { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({results:lastBatchResults}) });
    if (!res.ok) { showToast('Erreur lors de l\'export.', 'error'); return; }
    const blob  = await res.blob();
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement('a');
    a.href      = url;
    a.download  = `predictions_LD_${new Date().toISOString().slice(0,10)}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('Export Excel téléchargé', 'success');
  } catch(e) { showToast('Erreur réseau lors de l\'export', 'error'); }
}

// ── Update ratios ─────────────────────────────────────────────────────────────
async function loadUpdateStatus() {
  const bar = document.getElementById('updateStatus');
  if (!bar) return;
  try {
    const d = await fetch('/api/update-status').then(r => r.json());
    const dot = document.getElementById('updateDot');
    const ssDot = document.getElementById('sidebarStatusDot');
    const ssLbl = document.getElementById('sidebarStatusLbl');
    if (d.last_update) {
      bar.className = 'update-status-bar update-status-ok';
      bar.innerHTML = `<span class="us-dot us-dot-ok"></span>
        Dernière màj : <strong>${d.last_update}</strong>
        &nbsp;·&nbsp; Fichier : <code>${d.filename}</code>
        &nbsp;·&nbsp; Période : ${d.period_start} → ${d.period_end}
        &nbsp;·&nbsp; ${d.updated_lanes.length} lane(s) mises à jour`;
      if (dot) dot.style.display = 'block';
      if (ssDot) { ssDot.style.background = '#10B981'; }
      if (ssLbl) ssLbl.textContent = 'À jour';
    } else {
      bar.className = 'update-status-bar update-status-none';
      bar.innerHTML = `<span class="us-dot us-dot-none"></span>Aucune mise à jour — ratios initiaux actifs`;
      if (ssDot) { ssDot.style.background = '#F59E0B'; }
      if (ssLbl) ssLbl.textContent = 'Initial';
    }
  } catch(e) { if(bar) bar.textContent = 'Statut indisponible'; }
}
function handleMovesFileDrop(event) {
  event.preventDefault();
  document.getElementById('dropZoneUpdate').classList.remove('drag-over');
  const file = event.dataTransfer.files[0]; if (file) uploadMovesFile(file);
}
function handleMovesFile(event) { const f = event.target.files[0]; if (f) uploadMovesFile(f); event.target.value=''; }
async function uploadMovesFile(file) {
  const result = document.getElementById('updateResult');
  result.innerHTML = `<div style="padding:16px 0;font-size:13px;color:var(--muted)"><span class="spinner dark"></span>Traitement de "${file.name}"…</div>`;
  const fd = new FormData(); fd.append('file', file);
  try {
    const res  = await fetch('/api/update-ratios', { method:'POST', body:fd });
    const data = await res.json();
    if (!res.ok) { result.innerHTML = `<div class="error-box">❌ ${data.error}</div>`; showToast(data.error,'error'); return; }
    result.innerHTML = `
      <div class="update-report">
        <div class="update-report-header">✅ Ratios mis à jour
          <span class="update-report-meta">${data.total_rows.toLocaleString('fr-FR')} lignes · Période ${data.period_start} → ${data.period_end} · ${data.n_lanes} lanes actives</span>
        </div>
        ${data.updated_lanes.length ? `<div class="update-report-group"><div class="urg-label">✅ Mises à jour (${data.updated_lanes.length})</div><div class="urg-pills">${data.updated_lanes.map(l=>`<span class="pill pill-updated">${l}</span>`).join('')}</div></div>` : ''}
        ${data.new_lanes.length     ? `<div class="update-report-group"><div class="urg-label">🆕 Nouvelles (${data.new_lanes.length})</div><div class="urg-pills">${data.new_lanes.map(l=>`<span class="pill pill-new-lane">${l}</span>`).join('')}</div></div>` : ''}
        ${data.unchanged_lanes.length ? `<div class="update-report-group"><div class="urg-label">⏸ Non présentes (${data.unchanged_lanes.length})</div><div class="urg-pills">${data.unchanged_lanes.map(l=>`<span class="pill pill-unchanged">${l}</span>`).join('')}</div></div>` : ''}
      </div>`;
    showToast(`${data.n_lanes} lanes actives après mise à jour`, 'success');
    loadLanesDetails(); loadUpdateStatus(); loadDashboard();
  } catch(e) { result.innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`; showToast('Erreur réseau','error'); }
}

// ════════════════════════════════════════════════════════════════════
//  FONCTIONNALITÉ FULL / EMPTY
// ════════════════════════════════════════════════════════════════════

// ── Combobox FM (clone du combobox principal) ─────────────────────────────────
let selectedFMLane = '';

function buildComboListFM(lanes) {
  const list = document.getElementById('fmComboList');
  if (!list) return;
  list.innerHTML = '';
  if (!lanes.length) { list.innerHTML = '<div class="combo-empty">Aucune lane</div>'; return; }
  lanes.forEach(l => {
    const div = document.createElement('div');
    div.className = 'combo-option';
    div.innerHTML = `
      <span class="opt-dot" style="background:${freshColor(l.freshness)}"></span>
      <span class="opt-name">${l.lane}</span>
      <span class="opt-date">${l.last_date}</span>`;
    div.addEventListener('mousedown', e => { e.preventDefault(); selectFMLane(l); });
    list.appendChild(div);
  });
}
function toggleComboFM() {
  const cb = document.getElementById('fmLaneCombobox');
  const dd = document.getElementById('fmComboDropdown');
  const open = cb.classList.toggle('open');
  dd.style.display = open ? 'block' : 'none';
  if (open) { document.getElementById('fmComboSearch').focus(); filterComboFM(); }
}
function filterComboFM() {
  const q = document.getElementById('fmComboSearch').value.toLowerCase();
  buildComboListFM(allLanes.filter(l => l.lane.toLowerCase().includes(q)));
}
function selectFMLane(l) {
  selectedFMLane = l.lane;
  document.getElementById('fmSelectedLane').value    = l.lane;
  document.getElementById('fmComboText').textContent = l.lane;
  document.getElementById('fmComboDot').className    = `combo-dot ${l.freshness}`;
  document.getElementById('fmComboDropdown').style.display = 'none';
  document.getElementById('fmLaneCombobox').classList.remove('open');
  document.getElementById('fmComboSearch').value = '';
  document.getElementById('fmLaneText').value    = '';
}
document.addEventListener('click', e => {
  if (!document.getElementById('fmLaneCombobox')?.contains(e.target)) {
    if (document.getElementById('fmComboDropdown'))
      document.getElementById('fmComboDropdown').style.display = 'none';
    document.getElementById('fmLaneCombobox')?.classList.remove('open');
  }
});
function onFMLaneText() {
  const v = document.getElementById('fmLaneText').value.toUpperCase();
  document.getElementById('fmLaneText').value = v;
  if (v) {
    selectedFMLane = v;
    document.getElementById('fmComboText').textContent = '— Choisir —';
    document.getElementById('fmComboDot').className = 'combo-dot';
    document.getElementById('fmSelectedLane').value = '';
  }
}
function getFMLane() {
  return document.getElementById('fmSelectedLane').value ||
         document.getElementById('fmLaneText').value.trim().toUpperCase() || null;
}

// Hook: rebuild FM combobox when allLanes is loaded
const _origBuildComboList = buildComboList;
function buildComboList(lanes) { _origBuildComboList(lanes); buildComboListFM(lanes); }

// ── Predict FM ────────────────────────────────────────────────────────────────
async function predictFM() {
  const lane   = getFMLane();
  const volume = parseInt(document.getElementById('fmVolumeInput').value) || 0;
  if (!lane)   { showToast('Veuillez sélectionner ou saisir une lane.', 'warn'); return; }
  if (!volume) { showToast('Veuillez saisir un volume valide.', 'warn'); return; }

  const btn = document.getElementById('btnPredictFM');
  btn.textContent = 'Calcul…';
  setLoading('resultFM');

  try {
    const res  = await fetch('/api/predict/fm', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({lane, volume})
    });
    const data = await res.json();
    btn.textContent = 'Calculer D/L + Full & Empty';
    if (!res.ok) {
      document.getElementById('resultFM').innerHTML = `<div class="error-box">❌ ${data.error}</div>`;
      showToast(data.error, 'error'); return;
    }
    document.getElementById('resultFM').innerHTML = buildFMResultHTML(data);
    showToast(`Lane ${data.lane} — D-Full: ${fmt(data.D_full)} · L-Full: ${fmt(data.L_full)}`, 'success');
  } catch(e) {
    btn.textContent = 'Calculer D/L + Full & Empty';
    document.getElementById('resultFM').innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`;
    showToast('Erreur réseau', 'error');
  }
}

function buildFMResultHTML(d) {
  const circ    = 2 * Math.PI * 30;
  const loadArc = (d.pct_L / 100) * circ;
  const dischArc= (d.pct_D / 100) * circ;
  const warn    = d.outdated
    ? `<div class="warn-box">⚠ Dernière escale le <strong>${d.last_date}</strong> — données potentiellement obsolètes</div>` : '';
  const noFM    = !d.fm_available
    ? `<div class="warn-box">⚠ Ratios F/M non disponibles pour cette lane.</div>` : '';

  const fmBlock = d.fm_available ? `
    <div class="fm-section">
      <div class="fm-section-label">Étape 2 — Full / Empty par type de mouvement</div>
      <div class="fm-grid">
        <div class="fm-card d-full">
          <div class="fm-lbl">Discharge Full</div>
          <div class="fm-val">${fmt(d.D_full)}</div>
          <div class="fm-pct">${d.pct_D_full.toFixed(1)}% des Discharge</div>
        </div>
        <div class="fm-card d-empty">
          <div class="fm-lbl">Discharge Empty</div>
          <div class="fm-val">${fmt(d.D_empty)}</div>
          <div class="fm-pct">${d.pct_D_empty.toFixed(1)}% des Discharge</div>
        </div>
        <div class="fm-card l-full">
          <div class="fm-lbl">Load Full</div>
          <div class="fm-val">${fmt(d.L_full)}</div>
          <div class="fm-pct">${d.pct_L_full.toFixed(1)}% des Load</div>
        </div>
        <div class="fm-card l-empty">
          <div class="fm-lbl">Load Empty</div>
          <div class="fm-val">${fmt(d.L_empty)}</div>
          <div class="fm-pct">${d.pct_L_empty.toFixed(1)}% des Load</div>
        </div>
      </div>
    </div>` : noFM;

  return `
    <div class="result-section">
      <div class="result-header">Lane ${d.lane} &nbsp;·&nbsp; Volume ${fmt(d.volume)}</div>
      ${warn}
      <div style="font-family:var(--mono);font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">Étape 1 — Load / Discharge</div>
      <div class="result-grid" style="margin-bottom:16px">
        <div class="result-card load">
          <div class="rc-label">Load estimé</div>
          <div class="rc-value">${fmt(d.pred_L)}</div>
          <div class="rc-pct">${d.pct_L.toFixed(1)}% du volume</div>
        </div>
        <div class="result-card disch">
          <div class="rc-label">Discharge estimé</div>
          <div class="rc-value">${fmt(d.pred_D)}</div>
          <div class="rc-pct">${d.pct_D.toFixed(1)}% du volume</div>
        </div>
      </div>
      <div class="donut-wrap">
        <svg class="donut-svg" viewBox="0 0 80 80">
          <circle class="donut-track" cx="40" cy="40" r="30"/>
          <circle class="donut-load" cx="40" cy="40" r="30"
            stroke-dasharray="${loadArc} ${circ}" stroke-dashoffset="0" transform="rotate(-90 40 40)"/>
          <circle class="donut-disch" cx="40" cy="40" r="30"
            stroke-dasharray="${dischArc} ${circ}" stroke-dashoffset="${-loadArc}" transform="rotate(-90 40 40)"/>
          <text class="donut-text" x="40" y="37">L / D</text>
          <text class="donut-text" x="40" y="50" style="font-size:10px;fill:var(--ink);font-weight:500">${d.pct_L.toFixed(0)}/${d.pct_D.toFixed(0)}</text>
        </svg>
        <div class="donut-legend">
          <div class="donut-legend-item"><span class="dl-dot" style="background:var(--blue)"></span>Load<span class="dl-val" style="color:var(--blue)">${d.pct_L.toFixed(1)}%</span></div>
          <div class="donut-legend-item"><span class="dl-dot" style="background:var(--coral)"></span>Discharge<span class="dl-val" style="color:var(--coral)">${d.pct_D.toFixed(1)}%</span></div>
        </div>
      </div>
      ${fmBlock}
      <div class="info-grid" style="margin-top:12px">
        <div class="info-cell"><div class="ic-label">Escales analysées</div><div class="ic-value">${d.n_voyages} voyages</div></div>
        <div class="info-cell"><div class="ic-label">Période D/L</div><div class="ic-value">${d.period_start} → ${d.last_date}</div></div>
      </div>
    </div>`;
}

// ── Batch FM ──────────────────────────────────────────────────────────────────
let lastFMBatchResults = null;

function handleDropFM(event) {
  event.preventDefault();
  document.getElementById('dropZoneFM').classList.remove('drag-over');
  const f = event.dataTransfer.files[0]; if (f) uploadFileFM(f);
}
function handleFileFM(event) { const f = event.target.files[0]; if (f) uploadFileFM(f); event.target.value=''; }

async function uploadFileFM(file) {
  setLoading('batchFMResult', `Traitement D/L + F/M de "${file.name}"…`);
  const fd = new FormData(); fd.append('file', file);
  try {
    const res  = await fetch('/api/predict/fm/batch', {method:'POST', body:fd});
    const data = await res.json();
    if (!res.ok) { document.getElementById('batchFMResult').innerHTML = `<div class="warn-box">${data.error}</div>`; return; }
    renderFMBatch(data);
    showToast(`${data.count} lignes traitées (D/L + F/M)`, 'success');
  } catch(e) {
    document.getElementById('batchFMResult').innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`;
  }
}

function renderFMBatch(data) {
  lastFMBatchResults = data.results;
  const {results, errors, count} = data;
  const errHtml = errors.length
    ? `<div class="warn-box" style="margin-bottom:12px">${errors.length} ligne(s) ignorée(s) : ${errors.slice(0,3).join(' | ')}</div>` : '';

  const rows = results.map(r => {
    if (r.unknown) return `<tr>
      <td class="td-lane">${r.lane}</td><td>${fmt(r.volume)}</td>
      <td colspan="8"><span class="pill pill-unknown">Lane inconnue</span></td></tr>`;
    const wi  = r.outdated ? ' ⚠' : '';
    const fmOk = r.fm_available;
    return `<tr>
      <td class="td-lane">${r.lane}${wi}</td>
      <td style="font-family:var(--mono)">${fmt(r.volume)}</td>
      <td><span class="pill pill-d">${fmt(r.pred_D)}</span></td>
      <td><span class="pill pill-l">${fmt(r.pred_L)}</span></td>
      <td>${fmOk ? `<span class="pill pill-df">${fmt(r.D_full)}</span>` : '—'}</td>
      <td>${fmOk ? `<span class="pill pill-de">${fmt(r.D_empty)}</span>` : '—'}</td>
      <td>${fmOk ? `<span class="pill pill-lf">${fmt(r.L_full)}</span>` : '—'}</td>
      <td>${fmOk ? `<span class="pill pill-le">${fmt(r.L_empty)}</span>` : '—'}</td>
      <td style="font-family:var(--mono);font-size:11px;color:var(--muted)">${r.period_start||''} → ${r.last_date||''}</td>
    </tr>`;
  }).join('');

  document.getElementById('batchFMResult').innerHTML = `
    ${errHtml}
    <div class="batch-meta">
      <span>${count} ligne(s) traitée(s)</span>
      <button class="btn-secondary" onclick="exportFMResults()">Exporter .xlsx</button>
    </div>
    <div class="table-wrap"><table>
      <thead><tr>
        <th>Lane</th><th>Volume</th><th>Discharge</th><th>Load</th>
        <th>D-Full</th><th>D-Empty</th><th>L-Full</th><th>L-Empty</th><th>Période</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table></div>`;
}

async function exportFMResults() {
  if (!lastFMBatchResults) return;
  try {
    const res  = await fetch('/api/export/fm', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({results:lastFMBatchResults})});
    if (!res.ok) { showToast("Erreur lors de l'export.", 'error'); return; }
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = `predictions_FM_${new Date().toISOString().slice(0,10)}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('Export Excel F/M téléchargé', 'success');
  } catch(e) { showToast("Erreur réseau lors de l'export", 'error'); }
}

// ── Update FM status ──────────────────────────────────────────────────────────
async function loadFMUpdateStatus() {
  const bar = document.getElementById('fmUpdateStatus');
  if (!bar) return;
  try {
    const d = await fetch('/api/update-fm-status').then(r => r.json());
    if (d.last_update) {
      bar.className = 'update-status-bar update-status-ok';
      bar.innerHTML = `<span class="us-dot us-dot-ok"></span>
        Dernière màj F/M : <strong>${d.last_update}</strong>
        &nbsp;·&nbsp; <code>${d.filename}</code>
        &nbsp;·&nbsp; Période : ${d.period_start} → ${d.period_end}
        &nbsp;·&nbsp; ${d.n_lanes} lanes`;
    } else {
      bar.className = 'update-status-bar update-status-none';
      bar.innerHTML = `<span class="us-dot us-dot-none"></span>Aucune mise à jour F/M — ratios initiaux actifs`;
    }
  } catch(e) { if (bar) bar.textContent = 'Statut F/M indisponible'; }
}

function handleFMMovesFileDrop(event) {
  event.preventDefault();
  document.getElementById('dropZoneFMUpdate').classList.remove('drag-over');
  const f = event.dataTransfer.files[0]; if (f) uploadMovesFileFM(f);
}
function handleMovesFileFM(event) { const f = event.target.files[0]; if (f) uploadMovesFileFM(f); event.target.value=''; }

async function uploadMovesFileFM(file) {
  const result = document.getElementById('updateFMResult');
  result.innerHTML = `<div style="padding:16px 0;font-size:13px;color:var(--muted)"><span class="spinner dark"></span>Recalcul ratios F/M depuis "${file.name}"…</div>`;
  const fd = new FormData(); fd.append('file', file);
  try {
    const res  = await fetch('/api/update-fm-ratios', {method:'POST', body:fd});
    const data = await res.json();
    if (!res.ok) { result.innerHTML = `<div class="error-box">❌ ${data.error}</div>`; showToast(data.error,'error'); return; }
    result.innerHTML = `
      <div class="update-report">
        <div class="update-report-header">✅ Ratios F/M mis à jour
          <span class="update-report-meta">${data.total_rows.toLocaleString('fr-FR')} lignes · ${data.period_start} → ${data.period_end} · ${data.n_lanes} lanes</span>
        </div>
        ${data.updated_lanes.length ? `<div class="update-report-group"><div class="urg-label">✅ Mises à jour (${data.updated_lanes.length})</div><div class="urg-pills">${data.updated_lanes.map(l=>`<span class="pill pill-updated">${l}</span>`).join('')}</div></div>` : ''}
        ${data.new_lanes.length     ? `<div class="update-report-group"><div class="urg-label">🆕 Nouvelles (${data.new_lanes.length})</div><div class="urg-pills">${data.new_lanes.map(l=>`<span class="pill pill-new-lane">${l}</span>`).join('')}</div></div>` : ''}
      </div>`;
    showToast(`Ratios F/M mis à jour — ${data.n_lanes} lanes`, 'success');
    loadFMUpdateStatus();
  } catch(e) { result.innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`; showToast('Erreur réseau','error'); }
}

// Init FM au chargement
document.addEventListener('DOMContentLoaded', () => { loadFMUpdateStatus(); });

// ════════════════════════════════════════════════════════════════════
//  FONCTIONNALITÉ TEU (pipeline complet 3 étapes)
// ════════════════════════════════════════════════════════════════════

// ── Combobox TEU ──────────────────────────────────────────────────────────────
let selectedTEULane = '';

function buildComboListTEU(lanes) {
  const list = document.getElementById('teuComboList');
  if (!list) return;
  list.innerHTML = '';
  if (!lanes.length) { list.innerHTML = '<div class="combo-empty">Aucune lane</div>'; return; }
  lanes.forEach(l => {
    const div = document.createElement('div');
    div.className = 'combo-option';
    div.innerHTML = `
      <span class="opt-dot" style="background:${freshColor(l.freshness)}"></span>
      <span class="opt-name">${l.lane}</span>
      <span class="opt-date">${l.last_date}</span>`;
    div.addEventListener('mousedown', e => { e.preventDefault(); selectTEULane(l); });
    list.appendChild(div);
  });
}
function toggleComboTEU() {
  const cb = document.getElementById('teuLaneCombobox');
  const dd = document.getElementById('teuComboDropdown');
  const open = cb.classList.toggle('open');
  dd.style.display = open ? 'block' : 'none';
  if (open) { document.getElementById('teuComboSearch').focus(); filterComboTEU(); }
}
function filterComboTEU() {
  const q = document.getElementById('teuComboSearch').value.toLowerCase();
  buildComboListTEU(allLanes.filter(l => l.lane.toLowerCase().includes(q)));
}
function selectTEULane(l) {
  selectedTEULane = l.lane;
  document.getElementById('teuSelectedLane').value    = l.lane;
  document.getElementById('teuComboText').textContent = l.lane;
  document.getElementById('teuComboDot').className    = `combo-dot ${l.freshness}`;
  document.getElementById('teuComboDropdown').style.display = 'none';
  document.getElementById('teuLaneCombobox').classList.remove('open');
  document.getElementById('teuComboSearch').value = '';
  document.getElementById('teuLaneText').value    = '';
}
document.addEventListener('click', e => {
  if (!document.getElementById('teuLaneCombobox')?.contains(e.target)) {
    if (document.getElementById('teuComboDropdown'))
      document.getElementById('teuComboDropdown').style.display = 'none';
    document.getElementById('teuLaneCombobox')?.classList.remove('open');
  }
});
function onTEULaneText() {
  const v = document.getElementById('teuLaneText').value.toUpperCase();
  document.getElementById('teuLaneText').value = v;
  if (v) {
    selectedTEULane = v;
    document.getElementById('teuComboText').textContent = '— Choisir —';
    document.getElementById('teuComboDot').className = 'combo-dot';
    document.getElementById('teuSelectedLane').value = '';
  }
}
function getTEULane() {
  return document.getElementById('teuSelectedLane').value ||
         document.getElementById('teuLaneText').value.trim().toUpperCase() || null;
}

// Hook : rebuild TEU combobox when allLanes loads
const _origBuildComboListFM = typeof buildComboListFM !== 'undefined' ? buildComboListFM : null;
const _origBuildComboListForTEU = buildComboList;
function buildComboList(lanes) {
  _origBuildComboListForTEU(lanes);
  buildComboListTEU(lanes);
}

// ── Predict TEU ───────────────────────────────────────────────────────────────
async function predictTEU() {
  const lane   = getTEULane();
  const volume = parseInt(document.getElementById('teuVolumeInput').value) || 0;
  if (!lane)   { showToast('Veuillez sélectionner ou saisir une lane.', 'warn'); return; }
  if (!volume) { showToast('Veuillez saisir un volume valide.', 'warn'); return; }

  const btn = document.getElementById('btnPredictTEU');
  btn.textContent = 'Calcul…';
  setLoading('resultTEU');

  try {
    const res  = await fetch('/api/predict/teu', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({lane, volume})
    });
    const data = await res.json();
    btn.textContent = 'Calculer D/L + F/M + TEU';
    if (!res.ok) {
      document.getElementById('resultTEU').innerHTML = `<div class="error-box">❌ ${data.error}</div>`;
      showToast(data.error, 'error'); return;
    }
    document.getElementById('resultTEU').innerHTML = buildTEUResultHTML(data);
    showToast(`Lane ${data.lane} — Total TEU: ${fmt(data.total_teu)}`, 'success');
  } catch(e) {
    btn.textContent = 'Calculer D/L + F/M + TEU';
    document.getElementById('resultTEU').innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`;
    showToast('Erreur réseau', 'error');
  }
}

function buildTEUResultHTML(d) {
  const circ     = 2 * Math.PI * 30;
  const loadArc  = (d.pct_L / 100) * circ;
  const dischArc = (d.pct_D / 100) * circ;
  const warn     = d.outdated
    ? `<div class="warn-box">⚠ Dernière escale le <strong>${d.last_date}</strong> — données potentiellement obsolètes</div>` : '';

  // Étape 1 — D/L
  const step1 = `
    <div style="font-family:var(--mono);font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px">Étape 1 — Load / Discharge</div>
    <div class="result-grid" style="margin-bottom:16px">
      <div class="result-card load">
        <div class="rc-label">Load estimé</div>
        <div class="rc-value">${fmt(d.pred_L)}</div>
        <div class="rc-pct">${d.pct_L.toFixed(1)}% du volume</div>
      </div>
      <div class="result-card disch">
        <div class="rc-label">Discharge estimé</div>
        <div class="rc-value">${fmt(d.pred_D)}</div>
        <div class="rc-pct">${d.pct_D.toFixed(1)}% du volume</div>
      </div>
    </div>
    <div class="donut-wrap">
      <svg class="donut-svg" viewBox="0 0 80 80">
        <circle class="donut-track" cx="40" cy="40" r="30"/>
        <circle class="donut-load" cx="40" cy="40" r="30"
          stroke-dasharray="${loadArc} ${circ}" stroke-dashoffset="0" transform="rotate(-90 40 40)"/>
        <circle class="donut-disch" cx="40" cy="40" r="30"
          stroke-dasharray="${dischArc} ${circ}" stroke-dashoffset="${-loadArc}" transform="rotate(-90 40 40)"/>
        <text class="donut-text" x="40" y="37">L / D</text>
        <text class="donut-text" x="40" y="50" style="font-size:10px;fill:var(--ink);font-weight:500">${d.pct_L.toFixed(0)}/${d.pct_D.toFixed(0)}</text>
      </svg>
      <div class="donut-legend">
        <div class="donut-legend-item"><span class="dl-dot" style="background:var(--blue)"></span>Load<span class="dl-val" style="color:var(--blue)">${d.pct_L.toFixed(1)}%</span></div>
        <div class="donut-legend-item"><span class="dl-dot" style="background:var(--coral)"></span>Discharge<span class="dl-val" style="color:var(--coral)">${d.pct_D.toFixed(1)}%</span></div>
      </div>
    </div>`;

  // Étape 2 — F/M containers
  const step2 = d.fm_available ? `
    <div class="fm-section">
      <div class="fm-section-label">Étape 2 — Full / Empty (containers)</div>
      <div class="fm-grid">
        <div class="fm-card d-full"><div class="fm-lbl">Discharge Full</div><div class="fm-val">${fmt(d.D_full)}</div><div class="fm-pct">${(d.pct_D_full||0).toFixed(1)}% des D</div></div>
        <div class="fm-card d-empty"><div class="fm-lbl">Discharge Empty</div><div class="fm-val">${fmt(d.D_empty)}</div><div class="fm-pct">${(d.pct_D_empty||0).toFixed(1)}% des D</div></div>
        <div class="fm-card l-full"><div class="fm-lbl">Load Full</div><div class="fm-val">${fmt(d.L_full)}</div><div class="fm-pct">${(d.pct_L_full||0).toFixed(1)}% des L</div></div>
        <div class="fm-card l-empty"><div class="fm-lbl">Load Empty</div><div class="fm-val">${fmt(d.L_empty)}</div><div class="fm-pct">${(d.pct_L_empty||0).toFixed(1)}% des L</div></div>
      </div>
    </div>` : `<div class="warn-box" style="margin-top:14px">⚠ Ratios F/M non disponibles pour cette lane.</div>`;

  // Étape 3 — TEU
  const step3 = d.teu_available ? `
    <div class="teu-section">
      <div class="teu-section-label">Étape 3 — TEU (20' × 1 + 40' × 2)</div>
      <div class="teu-total-card">
        <div>
          <div class="teu-total-lbl">Total TEU estimé</div>
          <div style="font-size:11px;color:var(--muted);font-family:var(--mono);margin-top:2px">L-TEU: ${fmt(d.L_total_teu)} + D-TEU: ${fmt(d.D_total_teu)}</div>
        </div>
        <div class="teu-total-val">${fmt(d.total_teu)}</div>
      </div>
      <div class="teu-grid">
        <div class="teu-card d-full">
          <div class="teu-lbl">Discharge Full TEU</div>
          <div class="teu-containers">${fmt(d.D_full)} containers × ${d.ratio_D_full} TEU/cntr</div>
          <div class="teu-val">${fmt(d.D_full_teu)}</div>
        </div>
        <div class="teu-card d-empty">
          <div class="teu-lbl">Discharge Empty TEU</div>
          <div class="teu-containers">${fmt(d.D_empty)} containers × ${d.ratio_D_empty} TEU/cntr</div>
          <div class="teu-val">${fmt(d.D_empty_teu)}</div>
        </div>
        <div class="teu-card l-full">
          <div class="teu-lbl">Load Full TEU</div>
          <div class="teu-containers">${fmt(d.L_full)} containers × ${d.ratio_L_full} TEU/cntr</div>
          <div class="teu-val">${fmt(d.L_full_teu)}</div>
        </div>
        <div class="teu-card l-empty">
          <div class="teu-lbl">Load Empty TEU</div>
          <div class="teu-containers">${fmt(d.L_empty)} containers × ${d.ratio_L_empty} TEU/cntr</div>
          <div class="teu-val">${fmt(d.L_empty_teu)}</div>
        </div>
      </div>
    </div>` : `<div class="warn-box" style="margin-top:14px">⚠ Ratios TEU non disponibles pour cette lane.</div>`;

  return `
    <div class="result-section">
      <div class="result-header">Lane ${d.lane} &nbsp;·&nbsp; Volume ${fmt(d.volume)}</div>
      ${warn}
      ${step1}
      ${step2}
      ${step3}
    </div>`;
}

// ── Batch TEU ─────────────────────────────────────────────────────────────────
let lastTEUBatchResults = null;

function handleDropTEU(event) {
  event.preventDefault();
  document.getElementById('dropZoneTEU').classList.remove('drag-over');
  const f = event.dataTransfer.files[0]; if (f) uploadFileTEU(f);
}
function handleFileTEU(event) { const f = event.target.files[0]; if (f) uploadFileTEU(f); event.target.value=''; }

async function uploadFileTEU(file) {
  setLoading('batchTEUResult', `Traitement pipeline TEU complet de "${file.name}"…`);
  const fd = new FormData(); fd.append('file', file);
  try {
    const res  = await fetch('/api/predict/teu/batch', {method:'POST', body:fd});
    const data = await res.json();
    if (!res.ok) { document.getElementById('batchTEUResult').innerHTML = `<div class="warn-box">${data.error}</div>`; return; }
    renderTEUBatch(data);
    showToast(`${data.count} lignes traitées (D/L + F/M + TEU)`, 'success');
  } catch(e) {
    document.getElementById('batchTEUResult').innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`;
  }
}

function renderTEUBatch(data) {
  lastTEUBatchResults = data.results;
  const {results, errors, count} = data;
  const errHtml = errors.length
    ? `<div class="warn-box" style="margin-bottom:12px">${errors.length} ligne(s) ignorée(s) : ${errors.slice(0,3).join(' | ')}</div>` : '';

  const rows = results.map(r => {
    if (r.unknown) return `<tr>
      <td class="td-lane">${r.lane}</td><td>${fmt(r.volume)}</td>
      <td colspan="9"><span class="pill pill-unknown">Lane inconnue</span></td></tr>`;
    const wi   = r.outdated ? ' ⚠' : '';
    const tOk  = r.teu_available;
    return `<tr>
      <td class="td-lane">${r.lane}${wi}</td>
      <td style="font-family:var(--mono)">${fmt(r.volume)}</td>
      <td><span class="pill pill-d">${fmt(r.pred_D)}</span></td>
      <td><span class="pill pill-l">${fmt(r.pred_L)}</span></td>
      <td>${tOk ? `<span class="pill pill-dtf">${fmt(r.D_full_teu)}</span>` : '—'}</td>
      <td>${tOk ? `<span class="pill pill-dte">${fmt(r.D_empty_teu)}</span>` : '—'}</td>
      <td>${tOk ? `<span class="pill pill-ltf">${fmt(r.L_full_teu)}</span>` : '—'}</td>
      <td>${tOk ? `<span class="pill pill-lte">${fmt(r.L_empty_teu)}</span>` : '—'}</td>
      <td>${tOk ? `<span class="pill pill-teu">${fmt(r.total_teu)}</span>` : '—'}</td>
    </tr>`;
  }).join('');

  document.getElementById('batchTEUResult').innerHTML = `
    ${errHtml}
    <div class="batch-meta">
      <span>${count} ligne(s) traitée(s)</span>
      <button class="btn-secondary" onclick="exportTEUResults()">Exporter .xlsx</button>
    </div>
    <div class="table-wrap"><table>
      <thead><tr>
        <th>Lane</th><th>Volume</th><th>Discharge</th><th>Load</th>
        <th>D-Full TEU</th><th>D-Empty TEU</th><th>L-Full TEU</th><th>L-Empty TEU</th><th>Total TEU</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table></div>`;
}

async function exportTEUResults() {
  if (!lastTEUBatchResults) return;
  try {
    const res  = await fetch('/api/export/teu', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({results:lastTEUBatchResults})});
    if (!res.ok) { showToast("Erreur lors de l'export.", 'error'); return; }
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href = url; a.download = `predictions_TEU_${new Date().toISOString().slice(0,10)}.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showToast('Export TEU téléchargé', 'success');
  } catch(e) { showToast("Erreur réseau lors de l'export", 'error'); }
}

// ── Update TEU status ─────────────────────────────────────────────────────────
async function loadTEUUpdateStatus() {
  const bar = document.getElementById('teuUpdateStatus');
  if (!bar) return;
  try {
    const d = await fetch('/api/update-teu-status').then(r => r.json());
    if (d.last_update) {
      bar.className = 'update-status-bar update-status-ok';
      bar.innerHTML = `<span class="us-dot us-dot-ok"></span>
        Dernière màj TEU : <strong>${d.last_update}</strong>
        &nbsp;·&nbsp; <code>${d.filename}</code>
        &nbsp;·&nbsp; Période : ${d.period_start} → ${d.period_end}
        &nbsp;·&nbsp; ${d.n_lanes} lanes`;
    } else {
      bar.className = 'update-status-bar update-status-none';
      bar.innerHTML = `<span class="us-dot us-dot-none"></span>Aucune mise à jour TEU — ratios initiaux actifs`;
    }
  } catch(e) { if (bar) bar.textContent = 'Statut TEU indisponible'; }
}

function handleTEUMovesFileDrop(event) {
  event.preventDefault();
  document.getElementById('dropZoneTEUUpdate').classList.remove('drag-over');
  const f = event.dataTransfer.files[0]; if (f) uploadMovesFileTEU(f);
}
function handleMovesFileTEU(event) { const f = event.target.files[0]; if (f) uploadMovesFileTEU(f); event.target.value=''; }

async function uploadMovesFileTEU(file) {
  const result = document.getElementById('updateTEUResult');
  result.innerHTML = `<div style="padding:16px 0;font-size:13px;color:var(--muted)"><span class="spinner dark"></span>Recalcul ratios TEU depuis "${file.name}"…</div>`;
  const fd = new FormData(); fd.append('file', file);
  try {
    const res  = await fetch('/api/update-teu-ratios', {method:'POST', body:fd});
    const data = await res.json();
    if (!res.ok) { result.innerHTML = `<div class="error-box">❌ ${data.error}</div>`; showToast(data.error,'error'); return; }
    result.innerHTML = `
      <div class="update-report">
        <div class="update-report-header">✅ Ratios TEU mis à jour
          <span class="update-report-meta">${data.total_rows.toLocaleString('fr-FR')} lignes · ${data.period_start} → ${data.period_end} · ${data.n_lanes} lanes</span>
        </div>
        ${data.updated_lanes.length ? `<div class="update-report-group"><div class="urg-label">✅ Mises à jour (${data.updated_lanes.length})</div><div class="urg-pills">${data.updated_lanes.map(l=>`<span class="pill pill-updated">${l}</span>`).join('')}</div></div>` : ''}
        ${data.new_lanes.length     ? `<div class="update-report-group"><div class="urg-label">🆕 Nouvelles (${data.new_lanes.length})</div><div class="urg-pills">${data.new_lanes.map(l=>`<span class="pill pill-new-lane">${l}</span>`).join('')}</div></div>` : ''}
      </div>`;
    showToast(`Ratios TEU mis à jour — ${data.n_lanes} lanes`, 'success');
    loadTEUUpdateStatus();
  } catch(e) { result.innerHTML = `<div class="error-box">Erreur réseau : ${e.message}</div>`; showToast('Erreur réseau','error'); }
}

// Init TEU au chargement
document.addEventListener('DOMContentLoaded', () => { loadTEUUpdateStatus(); });
