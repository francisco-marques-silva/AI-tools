// ── Report tab ────────────────────────────────────────────────────────────────

const reportState = { files: [] };

const reportDropzone      = document.getElementById('reportDropzone');
const reportFileInput     = document.getElementById('reportFileInput');
const reportFileList      = document.getElementById('reportFileList');
const btnGenReport        = document.getElementById('btnGenReport');
const btnClearReport      = document.getElementById('btnClearReport');
const reportError         = document.getElementById('reportError');
const reportProgressCard  = document.getElementById('reportProgressCard');
const reportLog           = document.getElementById('reportLog');
const reportDownloads     = document.getElementById('reportDownloads');
const reportDownloadLinks = document.getElementById('reportDownloadLinks');

function detectFileType(name) {
  const n = name.toLowerCase();
  if (n === 'metadata.xlsx' || n === 'metadata.xls') return { label: 'Metadata', cls: '' };
  if (/^\d{8}\s*-/.test(name)) return { label: 'AI Result', cls: 'warn' };
  if (/-\s*tiab\.\w+$/i.test(name)) return { label: 'Human TIAB', cls: 'ok' };
  if (/-\s*fulltext\.\w+$/i.test(name)) return { label: 'Fulltext', cls: 'ok' };
  if (/-\s*listfinal\.\w+$/i.test(name)) return { label: 'Listfinal', cls: 'ok' };
  return { label: 'File', cls: '' };
}

function renderReportFileList() {
  if (!reportState.files.length) { reportFileList.classList.add('hidden'); return; }
  reportFileList.classList.remove('hidden');
  reportFileList.innerHTML = reportState.files.map((f, i) => {
    const type = detectFileType(f.name);
    const size = (f.size / 1024).toFixed(1) + ' KB';
    return `<div class="report-file-item">
      <span class="report-file-name">${f.name}</span>
      <span class="badge ${type.cls}">${type.label}</span>
      <span class="report-file-size">${size}</span>
      <button class="icon-btn remove-report-file" data-idx="${i}" title="Remove">×</button>
    </div>`;
  }).join('');
  reportFileList.querySelectorAll('.remove-report-file').forEach(btn => {
    btn.addEventListener('click', () => {
      reportState.files.splice(Number(btn.dataset.idx), 1);
      renderReportFileList();
      btnGenReport.disabled = !reportState.files.length;
    });
  });
  btnGenReport.disabled = false;
}

function addReportFiles(fileList) {
  for (const f of fileList) {
    if (!reportState.files.find(x => x.name === f.name)) reportState.files.push(f);
  }
  renderReportFileList();
}

reportDropzone.addEventListener('dragover', e => { e.preventDefault(); reportDropzone.classList.add('hover'); });
reportDropzone.addEventListener('dragleave', () => reportDropzone.classList.remove('hover'));
reportDropzone.addEventListener('drop', e => { e.preventDefault(); reportDropzone.classList.remove('hover'); if (e.dataTransfer.files.length) addReportFiles(e.dataTransfer.files); });
reportFileInput.addEventListener('change', e => { if (e.target.files.length) { addReportFiles(e.target.files); e.target.value = ''; } });

btnClearReport.addEventListener('click', () => {
  reportState.files = [];
  renderReportFileList();
  btnGenReport.disabled = true;
  reportProgressCard.classList.add('hidden');
  reportDownloads.classList.add('hidden');
  reportError.classList.add('hidden');
  document.getElementById('reportResultsCard').classList.add('hidden');
  _currentReportJobId = null;
});

// ── Chart gallery ──────────────────────────────────────────────────────────────

const CHART_LABELS = {
  'sensitivity_per_model_per_project':   'Sensitivity by Model per Project',
  'listfinal_capture_heatmap':           'Listfinal Capture Rate Heatmap',
  'test_retest_kappa':                   'Test-Retest Kappa with 95% CI',
  'model_comparison_radar':              'Model Comparison Radar',
  'cost_vs_sensitivity_bubble':          'Cost vs Sensitivity',
  'workload_reduction':                  'Workload Reduction',
  'efficiency_frontier_grid':            'Efficiency Frontier (Individual Runs)',
  'efficiency_frontier_by_project_avg':  'Efficiency Frontier (by Project)',
  'efficiency_frontier_averaged':        'Efficiency Frontier (Overall)',
  'efficiency_score_by_project':         'Efficiency Score by Project',
  'efficiency_score_aggregated':         'Efficiency Score Aggregated',
  'sensitivity_specificity_dual_gold':   'Sensitivity & Specificity (Dual Gold Standard)',
  'aggregated_performance_metrics':      'Aggregated Performance Metrics',
  'f1_score_vs_cost':                    'F1 Score vs Cost',
  'sensitivity_specificity_tradeoff':    'Sensitivity vs Specificity Trade-off',
  'model_ranking_heatmap':               'Model Ranking Heatmap',
};

function chartLabel(filename) {
  const key = filename.replace(/\.png$/, '');
  return CHART_LABELS[key] || key.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
}

function renderChartsGallery(jobId, chartFiles) {
  const grid = document.getElementById('chartsGrid');
  if (!chartFiles.length) { grid.innerHTML = '<p class="muted" style="padding:0 0 8px">No charts were generated.</p>'; return; }
  grid.innerHTML = chartFiles.map(fname => {
    const label = chartLabel(fname);
    const src = `/api/report/chart/${jobId}/${encodeURIComponent(fname)}`;
    return `<div class="chart-card" data-src="${src}" data-label="${label}">
      <img src="${src}" alt="${label}" loading="lazy" />
      <p class="chart-label">${label}</p>
    </div>`;
  }).join('');
  grid.querySelectorAll('.chart-card').forEach(card => {
    card.addEventListener('click', () => openLightbox(card.dataset.src, card.dataset.label));
  });
}

// ── Lightbox ──────────────────────────────────────────────────────────────────

const lightbox        = document.getElementById('chartLightbox');
const lightboxImg     = document.getElementById('lightboxImg');
const lightboxCaption = document.getElementById('lightboxCaption');

function openLightbox(src, caption) {
  lightboxImg.src = src;
  lightboxCaption.textContent = caption;
  lightbox.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
}

function closeLightbox() {
  lightbox.classList.add('hidden');
  lightboxImg.src = '';
  document.body.style.overflow = '';
}

document.querySelector('.lightbox-close').addEventListener('click', closeLightbox);
document.querySelector('.lightbox-backdrop').addEventListener('click', closeLightbox);
document.addEventListener('keydown', e => { if (e.key === 'Escape') closeLightbox(); });

// ── Table viewer ──────────────────────────────────────────────────────────────

let _currentReportJobId = null;

async function loadSheetList(jobId) {
  const sel = document.getElementById('sheetSelector');
  sel.innerHTML = '<option>Loading…</option>';
  try {
    const r = await fetch(`/api/report/tabledata/${jobId}`);
    if (!r.ok) { sel.innerHTML = '<option>Unavailable</option>'; return; }
    const data = await r.json();
    if (!data.sheets || !data.sheets.length) { sel.innerHTML = '<option>No sheets found</option>'; return; }
    sel.innerHTML = data.sheets.map(s => `<option value="${s}">${s}</option>`).join('');
    sel.addEventListener('change', () => loadSheetData(jobId, sel.value));
    loadSheetData(jobId, data.sheets[0]);
  } catch { sel.innerHTML = '<option>Error loading sheets</option>'; }
}

async function loadSheetData(jobId, sheet) {
  const wrap = document.getElementById('sheetTableWrap');
  wrap.innerHTML = '<p class="muted" style="padding:12px">Loading…</p>';
  try {
    const r = await fetch(`/api/report/tabledata/${jobId}/${encodeURIComponent(sheet)}`);
    if (!r.ok) { wrap.innerHTML = '<p class="muted" style="padding:12px">Failed to load sheet.</p>'; return; }
    const data = await r.json();
    if (!data.headers.length) { wrap.innerHTML = '<p class="muted" style="padding:12px">Sheet is empty.</p>'; return; }
    const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    wrap.innerHTML = `<table>
      <thead><tr>${data.headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
      <tbody>${data.rows.map(row => `<tr>${row.map(cell => `<td>${esc(cell)}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>`;
  } catch { wrap.innerHTML = '<p class="muted" style="padding:12px">Error loading data.</p>'; }
}

// ── Report generation ─────────────────────────────────────────────────────────

btnGenReport.addEventListener('click', async () => {
  if (!reportState.files.length) return;
  reportError.classList.add('hidden');
  reportProgressCard.classList.remove('hidden');
  reportLog.textContent = '';
  reportDownloads.classList.add('hidden');
  btnGenReport.disabled = true;

  const form = new FormData();
  reportState.files.forEach(f => form.append('files', f, f.name));

  let jobId;
  try {
    const resp = await fetch('/api/report/start', { method: 'POST', body: form });
    if (!resp.ok) {
      const txt = await resp.text().catch(() => '');
      reportError.textContent = `Failed to start report: ${resp.status}${txt ? ' — ' + txt : ''}`;
      reportError.classList.remove('hidden');
      btnGenReport.disabled = false;
      return;
    }
    const data = await resp.json();
    jobId = data.job_id;
  } catch (e) {
    reportError.textContent = String(e);
    reportError.classList.remove('hidden');
    btnGenReport.disabled = false;
    return;
  }

  const es = new EventSource(`/api/report/stream/${jobId}`);
  es.onmessage = ev => {
    try {
      const msg = JSON.parse(ev.data);
      if (msg.type === 'log') {
        reportLog.textContent += msg.line + '\n';
        reportLog.scrollTop = reportLog.scrollHeight;
      } else if (msg.type === 'done') {
        es.close();
        btnGenReport.disabled = false;
        if (msg.status === 'done') {
          const dlFiles = (msg.files || []).filter(f => !f.startsWith('charts'));
          if (dlFiles.length) {
            reportDownloadLinks.innerHTML = dlFiles.map(fname => {
              const icon = fname.endsWith('.docx')
                ? '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>'
                : '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>';
              return `<a class="btn btn-secondary" href="/api/report/download/${jobId}/${encodeURIComponent(fname)}" download="${fname}">${icon} ${fname}</a>`;
            }).join('');
            reportDownloads.classList.remove('hidden');
          }
          _currentReportJobId = jobId;
          const resultsCard = document.getElementById('reportResultsCard');
          resultsCard.classList.remove('hidden');
          renderChartsGallery(jobId, msg.charts || []);
          loadSheetList(jobId);
          const badge = document.getElementById('chartsCountBadge');
          if (badge && msg.charts && msg.charts.length) { badge.textContent = `${msg.charts.length} charts`; badge.classList.remove('hidden'); }
        } else if (msg.status === 'error') {
          reportError.textContent = 'Report generation failed. See log above for details.';
          reportError.classList.remove('hidden');
        }
      }
    } catch {}
  };
  es.onerror = () => {
    es.close();
    btnGenReport.disabled = false;
    reportError.textContent = 'Lost connection to server during report generation.';
    reportError.classList.remove('hidden');
  };
});
