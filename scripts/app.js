// ===== Tab switching =====
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    const target = btn.dataset.tab;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === target));
    document.querySelectorAll('.tab-pane').forEach(p => p.classList.toggle('hidden', p.id !== `pane-${target}`));
  });
});

// ===== State =====
const state = {
  workbook: null,
  filename: null,
  selectedSheet: null,
  rows: [],
};

// ===== Restore + persist preferences =====
document.addEventListener("DOMContentLoaded", () => {
  const savedModel = localStorage.getItem("modelSelect");
  const savedStudy = localStorage.getItem("studySynopsis");
  const savedInc = localStorage.getItem("inclusionCriteria");
  const savedExc = localStorage.getItem("exclusionCriteria");
  const savedKey = localStorage.getItem("apiKey");

  if (savedModel) document.getElementById("modelSelect").value = savedModel;
  if (savedStudy) {
    const el = document.getElementById("studySynopsis"); if (el) el.value = savedStudy;
  }
  if (savedInc) {
    const el = document.getElementById("inclusionCriteria"); if (el) el.value = savedInc;
  }
  if (savedExc) {
    const el = document.getElementById("exclusionCriteria"); if (el) el.value = savedExc;
  }
  if (savedKey) { const el=document.getElementById("apiKey"); if(el) el.value = savedKey; }

  // initialize parameter groups visibility
  const ms = document.getElementById("modelSelect");
  if (ms) {
    ms.addEventListener("change", toggleParamGroups);
    toggleParamGroups();
  }

  // Restore saved parameter values
  const paramIds = [
    "verbosity","reasoningEffort",
    "temperature",
  ];
  paramIds.forEach(id => {
    const v = localStorage.getItem(id);
    const el = document.getElementById(id);
    if (el && v != null) el.value = v;
  });
  // Attach tooltips for parameters
  addParamTooltips();

  initCriteriaLists();
  initAdvancedSettings();
});

const debounce = (fn, ms = 300) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };
const savePref = debounce((id, value) => localStorage.setItem(id, value), 200);
[
  "modelSelect",
  "studySynopsis",
  "inclusionCriteria",
  "exclusionCriteria",
  "apiKey",
  "verbosity","reasoningEffort",
  "temperature",
  "tierSelect",
  "adv_concurrent","adv_concurrent_max","adv_concurrent_min",
  "adv_record_retries","adv_aiup_after","adv_max_retries","adv_base_backoff",
].forEach((id) => {
  const el = document.getElementById(id);
  if (!el) return;
  const evt = el.tagName === "SELECT" ? "change" : "input";
  el.addEventListener(evt, () => savePref(id, el.value));
});

// ===== Tier Presets & Advanced Settings =====
const TIER_PRESETS = {
  free:  { concurrent: 2,  concurrent_max: 4,   concurrent_min: 1, record_retries: 3, aiup_after: 10, max_retries: 5, base_backoff: 2.0 },
  "1":   { concurrent: 5,  concurrent_max: 10,  concurrent_min: 1, record_retries: 3, aiup_after: 8,  max_retries: 5, base_backoff: 1.5 },
  "2":   { concurrent: 10, concurrent_max: 20,  concurrent_min: 2, record_retries: 3, aiup_after: 6,  max_retries: 5, base_backoff: 1.0 },
  "3":   { concurrent: 20, concurrent_max: 40,  concurrent_min: 2, record_retries: 3, aiup_after: 5,  max_retries: 5, base_backoff: 1.0 },
  "4":   { concurrent: 30, concurrent_max: 60,  concurrent_min: 3, record_retries: 3, aiup_after: 4,  max_retries: 5, base_backoff: 0.5 },
  "5":   { concurrent: 50, concurrent_max: 100, concurrent_min: 5, record_retries: 3, aiup_after: 3,  max_retries: 5, base_backoff: 0.5 },
};

function initAdvancedSettings() {
  const tierSel = document.getElementById('tierSelect');
  if (!tierSel) return;

  // Restore tier
  const savedTier = localStorage.getItem('tierSelect');
  if (savedTier) tierSel.value = savedTier;

  // Restore individual advanced fields
  const advIds = ['adv_concurrent','adv_concurrent_max','adv_concurrent_min',
    'adv_record_retries','adv_aiup_after','adv_max_retries','adv_base_backoff'];
  advIds.forEach(id => {
    const v = localStorage.getItem(id);
    const el = document.getElementById(id);
    if (el && v != null) el.value = v;
  });

  // If no saved values, apply current tier defaults
  const hasAny = advIds.some(id => localStorage.getItem(id) != null);
  if (!hasAny) applyTierPreset(tierSel.value || '3');

  tierSel.addEventListener('change', () => applyTierPreset(tierSel.value));
}

function applyTierPreset(tier) {
  const preset = TIER_PRESETS[tier];
  if (!preset) return;
  const map = {
    adv_concurrent: 'concurrent',
    adv_concurrent_max: 'concurrent_max',
    adv_concurrent_min: 'concurrent_min',
    adv_record_retries: 'record_retries',
    adv_aiup_after: 'aiup_after',
    adv_max_retries: 'max_retries',
    adv_base_backoff: 'base_backoff',
  };
  for (const [elId, key] of Object.entries(map)) {
    const el = document.getElementById(elId);
    if (el) {
      el.value = preset[key];
      savePref(elId, preset[key]);
    }
  }
}

function getAdvancedParams() {
  const num = (id) => {
    const el = document.getElementById(id);
    if (!el) return undefined;
    const v = el.value?.trim();
    if (v === '' || v == null) return undefined;
    const n = Number(v);
    return Number.isFinite(n) ? n : undefined;
  };
  const adv = {};
  const c = num('adv_concurrent');      if (c !== undefined) adv.concurrency = c;
  const cm = num('adv_concurrent_max'); if (cm !== undefined) adv.concurrency_max = cm;
  const cmin = num('adv_concurrent_min'); if (cmin !== undefined) adv.concurrent_min = cmin;
  const rr = num('adv_record_retries'); if (rr !== undefined) adv.record_max_retries = rr;
  const aiup = num('adv_aiup_after');   if (aiup !== undefined) adv.aiup_after = aiup;
  const mr = num('adv_max_retries');    if (mr !== undefined) adv.max_retries = mr;
  const bb = num('adv_base_backoff');   if (bb !== undefined) adv.base_backoff = bb;
  return adv;
}

// ===== Upload & read =====
const fileInput = document.getElementById("fileInput");
const dropzone = document.getElementById("dropzone");
const fileMeta = document.getElementById("fileMeta");
const sheetPickerWrap = document.getElementById("sheetPickerWrap");
const sheetSelect = document.getElementById("sheetSelect");
const columnsStatus = document.getElementById("columnsStatus");
const previewWrap = document.getElementById("previewWrap");
const previewTable = document.getElementById("previewTable");
const errorBanner = document.getElementById("errorBanner");
let lastPayload = null; // stores last built payload for restart

function showError(msg) { if (!errorBanner) { alert(msg); return; } errorBanner.textContent = msg; errorBanner.classList.remove("hidden"); }
function clearError() { if (!errorBanner) return; errorBanner.textContent = ""; errorBanner.classList.add("hidden"); }

dropzone.addEventListener("dragover", (e) => { e.preventDefault(); dropzone.classList.add("hover"); });
dropzone.addEventListener("dragleave", () => dropzone.classList.remove("hover"));
dropzone.addEventListener("drop", (e) => { e.preventDefault(); dropzone.classList.remove("hover"); const f = e.dataTransfer.files?.[0]; if (f) handleFile(f); });
fileInput.addEventListener("change", (e) => { const f = e.target.files?.[0]; if (f) handleFile(f); });

async function ensureXLSX() {
  if (typeof XLSX !== "undefined") return true;
  // Try to load from CDN dynamically (fallback if head script failed)
  const trySrc = (src) => new Promise((resolve, reject) => { const s=document.createElement('script'); s.src=src; s.defer=true; s.onload=resolve; s.onerror=reject; document.head.appendChild(s); });
  try {
    await trySrc('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js');
    if (typeof XLSX !== "undefined") return true;
  } catch {}
  try {
    await trySrc('https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js');
    if (typeof XLSX !== "undefined") return true;
  } catch {}
  return false;
}

async function handleFile(file) {
  state.filename = file.name;
  fileMeta.textContent = `${file.name} • ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  const reader = new FileReader();
  reader.onload = (evt) => {
    try {
      const data = evt.target.result;
      const lower = file.name.toLowerCase();
      if (lower.endsWith(".csv")) {
        const text = decodeTextWithFallback(data);
        const rows = csvToArray(text);
        if (!rows.length) showError("No rows could be parsed from the CSV. Check delimiter, encoding and header row.");
        state.workbook = null; state.rows = rows; sheetPickerWrap.classList.add("hidden"); validateAndPreview();
      } else if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        if (typeof XLSX === "undefined") {
          // Attempt dynamic load
          ensureXLSX().then(ok => {
            if (!ok) { showError("Excel library not loaded. Ensure internet or load the page via http://localhost:8000."); return; }
            // Re-run parsing after loading
            try {
              const wb = XLSX.read(data, { type: "array" });
              if (!wb.SheetNames?.length) { showError("No sheet found in this Excel file."); return; }
              state.workbook = wb; sheetSelect.innerHTML = "";
              wb.SheetNames.forEach((s, i) => { const opt = document.createElement("option"); opt.value = s; opt.textContent = s; if (i===0) opt.selected = true; sheetSelect.appendChild(opt); });
              sheetPickerWrap.classList.remove("hidden"); state.selectedSheet = wb.SheetNames[0]; loadSheet(state.selectedSheet);
            } catch (e) { showError(`Failed to parse Excel: ${e?.message || e}`); }
          });
          return;
        }
        const wb = XLSX.read(data, { type: "array" });
        if (!wb.SheetNames?.length) { showError("No sheet found in this Excel file."); return; }
        state.workbook = wb; sheetSelect.innerHTML = "";
        wb.SheetNames.forEach((s, i) => { const opt = document.createElement("option"); opt.value = s; opt.textContent = s; if (i===0) opt.selected = true; sheetSelect.appendChild(opt); });
        sheetPickerWrap.classList.remove("hidden"); state.selectedSheet = wb.SheetNames[0]; loadSheet(state.selectedSheet);
      } else {
        showError("Unsupported file type. Upload .csv, .xlsx, or .xls.");
      }
    } catch (e) { showError(`Failed to read file: ${e?.message || e}`); }
  };
  reader.readAsArrayBuffer(file);
}

sheetSelect.addEventListener("change", () => { state.selectedSheet = sheetSelect.value; loadSheet(state.selectedSheet); });

function loadSheet(name) {
  const ws = state.workbook.Sheets[name];
  const json = XLSX.utils.sheet_to_json(ws, { defval: "" });
  state.rows = json; validateAndPreview();
}

// ===== Criteria dynamic lists =====
function initCriteriaLists(){
  const incContainer = document.getElementById('inclusionList');
  const excContainer = document.getElementById('exclusionList');
  const addInc = document.getElementById('addInc');
  const addExc = document.getElementById('addExc');
  if (!incContainer || !excContainer) return;
  // seed from saved newline-separated values for compatibility
  const savedInc = localStorage.getItem('inclusionCriteria') || '';
  const savedExc = localStorage.getItem('exclusionCriteria') || '';
  const incArr = savedInc.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  const excArr = savedExc.split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
  if (!incArr.length) incArr.push('');
  if (!excArr.length) excArr.push('');
  incArr.forEach(v=>addCriteriaInput(incContainer, v, 'inc'));
  excArr.forEach(v=>addCriteriaInput(excContainer, v, 'exc'));
  if (addInc) addInc.onclick = () => addCriteriaInput(incContainer, '', 'inc');
  if (addExc) addExc.onclick = () => addCriteriaInput(excContainer, '', 'exc');
}
function addCriteriaInput(container, value, kind){
  const row = document.createElement('div');
  row.className = 'criteria-row';
  const input = document.createElement('input');
  input.type = 'text';
  input.placeholder = kind === 'inc' ? 'e.g., population: rats or mice (preclinical)' : 'e.g., case reports, reviews, editorials';
  input.value = value || '';
  input.addEventListener('input', persistCriteria);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      // add a new row of the same kind
      addCriteriaInput(container, '', kind);
    }
  });
  const remove = document.createElement('button');
  remove.type = 'button';
  remove.className = 'icon-btn remove-criterion';
  remove.title = 'Remove this criterion';
  remove.setAttribute('aria-label', 'Remove');
  remove.textContent = '×';
  remove.addEventListener('click', () => {
    const total = container.querySelectorAll('input').length;
    if (total > 1) {
      container.removeChild(row);
    } else {
      input.value = '';
    }
    persistCriteria();
  });
  row.appendChild(input);
  row.appendChild(remove);
  container.appendChild(row);
  // focus the new input for quick expansion UX
  setTimeout(() => input.focus(), 0);
}
function persistCriteria(){
  const incVals = Array.from(document.querySelectorAll('#inclusionList input')).map(el=>el.value.trim()).filter(Boolean);
  const excVals = Array.from(document.querySelectorAll('#exclusionList input')).map(el=>el.value.trim()).filter(Boolean);
  localStorage.setItem('inclusionCriteria', incVals.join('\n'));
  localStorage.setItem('exclusionCriteria', excVals.join('\n'));
}

// ===== Model parameter helpers =====
function isReasoningModel(model){ return /^gpt-5/.test(model || ""); }
function isChatModel(model){ return /^(gpt-4\.1|gpt-4o|gpt-3\.5)/.test(model || ""); }
function toggleParamGroups(){
  const model = (document.getElementById("modelSelect")?.value || "");
  const r = document.getElementById("paramsReasoning");
  const c = document.getElementById("paramsChat");
  if (!r || !c) return;
  if (isReasoningModel(model)) { r.classList.remove("hidden"); c.classList.add("hidden"); }
  else if (isChatModel(model)) { c.classList.remove("hidden"); r.classList.add("hidden"); }
  else { r.classList.add("hidden"); c.classList.add("hidden"); }
}

function validateParamsOrWarn(){
  const model = (document.getElementById("modelSelect")?.value || "");
  const warn = document.getElementById("paramsWarning");
  if (warn) { warn.classList.add("hidden"); warn.textContent = ""; }
  const num = (id) => {
    const el=document.getElementById(id); if(!el) return undefined; const v=el.value?.trim(); if(v===""||v==null) return undefined; const n=Number(v); return Number.isFinite(n)?n:NaN;
  };
  if (isReasoningModel(model)){
    // no numeric validation needed for reasoning controls
  } else if (isChatModel(model)){
    const t = num("temperature");
    if (t!==undefined && (t<0 || t>2)) { if(warn){ warn.textContent = "Temperature must be between 0 and 2."; warn.classList.remove("hidden"); } return false; }
  }
  return true;
}

function addParamTooltips(){
  const tips = {
    verbosity: 'Controls the level of detail in the response.\nlow: concise • medium: balanced (recommended) • high: very detailed.',
    reasoningEffort: 'Controls how much the model “thinks” before answering.\nminimal: fastest • low/medium: good balance (recommended) • high: most thorough but slower.',
    maxTokensReasoning: 'Upper bound on tokens generated for the answer. Use a positive integer. Typical: 256–1024.',
    temperature: 'Controls randomness: 0 = deterministic, 2 = very creative. Recommended: 0.2–0.8.',
    topP: 'Nucleus sampling: consider tokens whose cumulative probability ≤ top_p. Recommended: 0.8–1.0. Prefer tuning either temperature or top_p, not both.',
    presencePenalty: 'Encourages introducing new topics (higher = more encouragement). Recommended: 0–0.5.',
    frequencyPenalty: 'Discourages repetition of the same tokens (higher = less repetition). Recommended: 0–0.5.',
    maxTokensChat: 'Upper bound on tokens generated for the answer. Use a positive integer. Typical: 256–1024.',
  };
  Object.entries(tips).forEach(([id, text]) => {
    const el = document.getElementById(id);
    if (!el) return;
    const label = el.closest('label.field');
    if (!label) return;
    const titleSpan = label.querySelector('span');
    if (!titleSpan) return;
    if (titleSpan.querySelector('.help')) return; // avoid duplicates
    const help = document.createElement('span');
    help.className = 'help';
    help.tabIndex = 0;
    const icon = document.createElement('span'); icon.className = 'icon'; icon.textContent = '?';
    const tip = document.createElement('span'); tip.className = 'tip'; tip.textContent = text;
    help.appendChild(icon); help.appendChild(tip);
    titleSpan.appendChild(help);
  });
}

// ===== CSV helpers =====
function csvToArray(text) {
  const lines = text.split(/\r?\n/).filter((l) => l.length > 0);
  if (!lines.length) return [];
  const delim = detectDelimiter(lines[0]);
  if (lines[0].charCodeAt(0) === 0xfeff) lines[0] = lines[0].slice(1);
  const header = splitCsvLine(lines[0], delim).map((s) => s.trim());
  const out = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = splitCsvLine(lines[i], delim);
    const obj = {}; header.forEach((h, idx) => obj[h] = cols[idx] ?? "");
    out.push(obj);
  }
  return out;
}
function detectDelimiter(headerLine) { const commas = (headerLine.match(/,/g) || []).length; const semis = (headerLine.match(/;/g) || []).length; return semis > commas ? ";" : ","; }
function splitCsvLine(line, delim) {
  const out = []; let cur = ""; let inQ = false;
  for (let i=0;i<line.length;i++){ const c=line[i]; if(c==='"'){ if(inQ && line[i+1]==='"'){ cur+='"'; i++; } else { inQ=!inQ; } continue; } if(c===delim && !inQ){ out.push(cur); cur=""; continue;} cur+=c; }
  out.push(cur); return out.map((s)=>s.trim());
}
function decodeTextWithFallback(buffer){ const dec=(enc)=>new TextDecoder(enc,{fatal:false}).decode(buffer); let txt=dec("utf-8"); const bad=(txt.match(/\uFFFD/g)||[]).length; if(bad>0 && bad/Math.max(1,txt.length)>0.001){ try{ txt=dec("windows-1252"); }catch{} } return txt; }

// ===== Validation + preview =====
function validateAndPreview(){
  const cols = inferColumns(state.rows);
  const hasTitle = cols.includes("title"); const hasAbstract = cols.includes("abstract");
  let statusHtml = "";
  if (hasTitle && hasAbstract) statusHtml = `<strong style="color:var(--ok)">✓</strong> Detected required columns: <code>title</code>, <code>abstract</code>.`;
  else if (!state.rows.length) statusHtml = `No rows could be read from this sheet/file.`;
  else { const miss=[]; if(!hasTitle) miss.push("title"); if(!hasAbstract) miss.push("abstract"); statusHtml = `<span class="badge bad">Missing</span> Required columns not found: <code>${miss.join("</code>, <code>")}</code>.`; }
  columnsStatus.innerHTML = statusHtml;

  const head = previewTable.querySelector("thead"); const body = previewTable.querySelector("tbody"); head.innerHTML = ""; body.innerHTML = "";
  if (!state.rows.length) { previewWrap.classList.add("hidden"); return; }
  const colsToShow = cols.length ? cols : Object.keys(state.rows[0] ?? {});
  const trh=document.createElement("tr"); colsToShow.forEach(c=>{ const th=document.createElement("th"); th.textContent=c; trh.appendChild(th); }); head.appendChild(trh);
  const ell=(s,n=200)=>{ if(s==null) return ""; const str=String(s); return str.length>n?str.slice(0,n-1)+"…":str; };
  state.rows.slice(0,5).forEach(r=>{ const tr=document.createElement("tr"); colsToShow.forEach(c=>{ const td=document.createElement("td"); const full=r[c]??""; td.textContent=ell(full); if(String(full).length>200) td.title=String(full); tr.appendChild(td); }); body.appendChild(tr); });
  previewWrap.classList.remove("hidden");
  document.getElementById("btnSend").disabled = !(hasTitle && hasAbstract);
  if (state.filename) {
    const sizePart = (fileMeta.textContent.match(/•\s[\d\.]+\sMB/) || [""])[0];
    fileMeta.textContent = sizePart ? `${state.filename} ${sizePart} • ${state.rows.length.toLocaleString()} rows` : `${state.filename} • ${state.rows.length.toLocaleString()} rows`;
  }
}

function inferColumns(rows){ if(!rows.length) return []; const rawCols=Object.keys(rows[0]); const map=rawCols.reduce((a,c)=>{a[c.toLowerCase().trim()]=c; return a;},{}); const variants={ title:["title","título","titulo"], abstract:["abstract","resumo","summary"]}; const normalized=[]; const chosen={}; for(const k in variants){ for(const v of variants[k]){ if(map[v]){ chosen[k]=map[v]; break; } } } if(chosen.title||chosen.abstract){ for(const row of rows){ const r2={...row}; if(chosen.title) r2.title=row[chosen.title]; if(chosen.abstract) r2.abstract=row[chosen.abstract]; normalized.push(r2);} state.rows=normalized; return Object.keys(normalized[0]??{});} return rawCols; }

// ===== Payload =====
function buildPayload(){
  const model = (document.getElementById("modelSelect")?.value || "");
  const synopsis = (document.getElementById("studySynopsis")?.value || "");
  const incl = (document.getElementById("inclusionCriteria")?.value || "");
  const excl = (document.getElementById("exclusionCriteria")?.value || "");
  // Gather criteria from dynamic lists if present
  const incListEls = document.querySelectorAll('#inclusionList input');
  const excListEls = document.querySelectorAll('#exclusionList input');
  const inclusionArr = incListEls.length ? Array.from(incListEls).map(el=>el.value.trim()).filter(Boolean) : splitLines(incl);
  const exclusionArr = excListEls.length ? Array.from(excListEls).map(el=>el.value.trim()).filter(Boolean) : splitLines(excl);
  const api_key = (document.getElementById("apiKey").value || "").trim();
  if (!model) throw new Error("Select a model.");
  if (!state.rows.length) throw new Error("Load a spreadsheet.");
  if (!api_key) throw new Error("Enter your OpenAI API key.");
  if (!validateParamsOrWarn()) throw new Error("Please fix parameters.");
  const first = state.rows[0] ?? {};
  if (!("title" in first) || !("abstract" in first)) throw new Error("The spreadsheet must have 'title' and 'abstract' columns (or common variants).");
  // build params by family
  const params = {};
  if (isReasoningModel(model)){
    const v = document.getElementById("verbosity")?.value?.trim();
    const re = document.getElementById("reasoningEffort")?.value?.trim();
    if (v) params.verbosity = v;
    if (re) params.reasoning_effort = re;
  } else if (isChatModel(model)){
    const t = document.getElementById("temperature")?.value?.trim();
    if (t) params.temperature = Number(t);
  }
  // Merge advanced concurrency params
  const advParams = getAdvancedParams();
  Object.assign(params, advParams);
  return {
    model,
    api_key,
    study_synopsis: synopsis,
    inclusion_criteria: inclusionArr,
    exclusion_criteria: exclusionArr,
    params,
    sheet: state.selectedSheet || state.filename || "",
    filename: state.filename || "",
    sample_preview: state.rows.slice(0,5).map(r=>({ title: r.title ?? "", abstract: r.abstract ?? "" })),
    normalized_columns: true,
    records: state.rows.map((r,i)=>({ id: r.id ?? r.ID ?? r.Id ?? i+1, title: r.title ?? "", abstract: r.abstract ?? "" })),
  };
}
function splitLines(txt){ return txt.split(/\r?\n/).map(s=>s.replace(/^[-•\s]+/, "").trim()).filter(Boolean); }

// ===== Actions =====
const btnGenerate = document.getElementById("btnGenerate");
const btnSend = document.getElementById("btnSend");
const payloadCard = document.getElementById("payloadCard");
const payloadOut = document.getElementById("payloadOut");
const btnCopy = document.getElementById("btnCopy");
const btnClear = document.getElementById("btnClear");

btnGenerate.addEventListener("click", () => { try { const payload=buildPayload(); payloadOut.textContent = JSON.stringify(payload,null,2); payloadCard.classList.remove("hidden"); } catch(e){ showError(e.message); } });
btnCopy.addEventListener("click", async () => { try { await navigator.clipboard.writeText(payloadOut.textContent); btnCopy.textContent="Copied!"; setTimeout(()=>btnCopy.textContent="Copy",1200);} catch{ const ta=document.createElement("textarea"); ta.value=payloadOut.textContent; document.body.appendChild(ta); ta.select(); document.execCommand("copy"); document.body.removeChild(ta);} });

btnSend.addEventListener("click", async () => {
  try {
    clearError();
    const payload = buildPayload();
    lastPayload = payload;
    const prevHTML = btnSend.innerHTML; btnSend.textContent = "Sending…"; btnSend.disabled = true;
    const controller = new AbortController(); const timeout = setTimeout(()=>controller.abort(), 30000);
    const resp = await fetch(`/api/start`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload), signal: controller.signal }).finally(()=>clearTimeout(timeout));
    if (!resp.ok) {
      const txt = await resp.text().catch(()=>"");
      let friendly = `Send failed (${resp.status}).`;
      if (resp.status === 501) friendly += " The server does not support POST (likely a static server or file://). Run 'uvicorn backend:app --port 8000' and open http://localhost:8000.";
      else if (resp.status === 404) friendly += " Endpoint not found. Access via http://localhost:8000 and ensure /api/start exists.";
      else if (resp.status === 405) friendly += " Method not allowed. You're likely hitting a static server, not FastAPI.";
      showError(friendly + (txt ? `\nDetails: ${txt}` : ""));
      return;
    }
    const data = await resp.json(); const jobId = data.job_id; showProgress(jobId);
  } catch (e) { showError(e.message || String(e)); }
  finally { btnSend.innerHTML = prevHTML; btnSend.disabled = false; }
});

// Clear spreadsheet handler
if (btnClear){
  btnClear.addEventListener("click", () => {
    // Reset state
    state.workbook = null;
    state.filename = null;
    state.selectedSheet = null;
    state.rows = [];
    // Reset UI
    const fi = document.getElementById("fileInput"); if (fi) fi.value = "";
    fileMeta.textContent = "";
    sheetPickerWrap.classList.add("hidden");
    columnsStatus.textContent = "Waiting for file...";
    const head = previewTable.querySelector("thead"); const body = previewTable.querySelector("tbody"); if (head) head.innerHTML = ""; if (body) body.innerHTML = "";
    previewWrap.classList.add("hidden");
    payloadCard.classList.add("hidden");
    document.getElementById("btnSend").disabled = true;
    lastPayload = null;
  });
}

// ===== Report Generation =====
const reportState = { files: [] };

const reportDropzone = document.getElementById('reportDropzone');
const reportFileInput = document.getElementById('reportFileInput');
const reportFileList = document.getElementById('reportFileList');
const btnGenReport = document.getElementById('btnGenReport');
const btnClearReport = document.getElementById('btnClearReport');
const reportError = document.getElementById('reportError');
const reportProgressCard = document.getElementById('reportProgressCard');
const reportLog = document.getElementById('reportLog');
const reportDownloads = document.getElementById('reportDownloads');
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

// ── Results tab switching ──────────────────────────────────────────
document.querySelectorAll('.res-tab').forEach(btn => {
  btn.addEventListener('click', () => {
    const target = btn.dataset.rt;
    document.querySelectorAll('.res-tab').forEach(b => b.classList.toggle('active', b.dataset.rt === target));
    document.querySelectorAll('.res-pane').forEach(p => p.classList.toggle('hidden', p.id !== `rt-${target}`));
  });
});

// ── Chart gallery ─────────────────────────────────────────────────
const CHART_LABELS = {
  'sensitivity_per_model_per_project': 'Sensitivity by Model per Project',
  'listfinal_capture_heatmap': 'Listfinal Capture Rate Heatmap',
  'test_retest_kappa': 'Test-Retest Kappa with 95% CI',
  'model_comparison_radar': 'Model Comparison Radar',
  'cost_vs_sensitivity_bubble': 'Cost vs Sensitivity',
  'workload_reduction': 'Workload Reduction',
  'efficiency_frontier_grid': 'Efficiency Frontier (Individual Runs)',
  'efficiency_frontier_by_project_avg': 'Efficiency Frontier (by Project)',
  'efficiency_frontier_averaged': 'Efficiency Frontier (Overall)',
  'efficiency_score_by_project': 'Efficiency Score by Project',
  'efficiency_score_aggregated': 'Efficiency Score Aggregated',
  'sensitivity_specificity_dual_gold': 'Sensitivity & Specificity (Dual Gold Standard)',
  'aggregated_performance_metrics': 'Aggregated Performance Metrics',
  'f1_score_vs_cost': 'F1 Score vs Cost',
  'sensitivity_specificity_tradeoff': 'Sensitivity vs Specificity Trade-off',
  'model_ranking_heatmap': 'Model Ranking Heatmap',
};

function chartLabel(filename) {
  const key = filename.replace(/\.png$/, '');
  return CHART_LABELS[key] || key.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
}

function renderChartsGallery(jobId, chartFiles) {
  const grid = document.getElementById('chartsGrid');
  if (!chartFiles.length) { grid.innerHTML = '<p class="muted" style="padding:16px">No charts were generated.</p>'; return; }
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

// ── Lightbox ──────────────────────────────────────────────────────
const lightbox = document.getElementById('chartLightbox');
const lightboxImg = document.getElementById('lightboxImg');
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

// ── Table viewer ──────────────────────────────────────────────────
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
    const esc = s => String(s ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    wrap.innerHTML = `<table>
      <thead><tr>${data.headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
      <tbody>${data.rows.map(row => `<tr>${row.map(cell => `<td>${esc(cell)}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>`;
  } catch { wrap.innerHTML = '<p class="muted" style="padding:12px">Error loading data.</p>'; }
}

// ── Report generation ──────────────────────────────────────────────
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
          // download buttons
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
          // charts + tables results pane
          _currentReportJobId = jobId;
          const resultsCard = document.getElementById('reportResultsCard');
          resultsCard.classList.remove('hidden');
          renderChartsGallery(jobId, msg.charts || []);
          loadSheetList(jobId);
          // update charts tab label
          const chartsTab = document.querySelector('[data-rt="charts"]');
          if (chartsTab && msg.charts && msg.charts.length) chartsTab.innerHTML = chartsTab.innerHTML.replace('Charts', `Charts (${msg.charts.length})`);
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

// ===== Progress (SSE) + cancel =====
function showProgress(jobId){
  const card = document.getElementById("progressCard"); const bar = document.getElementById("progressBar"); const label = document.getElementById("progressLabel"); const link = document.getElementById("downloadLink"); const linkX = document.getElementById("downloadLinkXlsx"); const btnCancel = document.getElementById("btnCancel");
  const btnRestart = document.getElementById("btnRestart");
  const liveLog = document.getElementById("liveLog");
  const btnToggleLog = document.getElementById("btnToggleLog");
  card.classList.remove("hidden"); link.classList.add("hidden"); if(linkX) linkX.classList.add("hidden"); bar.style.width = "0%"; label.textContent = "Starting...";
  // Reset log and keep hidden by default to avoid UI pollution
  if (liveLog) { liveLog.textContent = ""; liveLog.classList.add("hidden"); }
  if (btnToggleLog) {
    btnToggleLog.textContent = "Show live log";
    btnToggleLog.onclick = () => {
      if (!liveLog) return;
      const hidden = liveLog.classList.contains("hidden");
      if (hidden) {
        liveLog.classList.remove("hidden");
        btnToggleLog.textContent = "Hide live log";
      } else {
        liveLog.classList.add("hidden");
        btnToggleLog.textContent = "Show live log";
      }
    };
  }

  // polling state for partial results
  let pollHandle = null;
  let delivered = 0;
  function startPolling(){
    if (pollHandle) return;
    pollHandle = setInterval(async () => {
      try {
        const resp = await fetch(`/api/partial/${jobId}?since=${delivered}`);
        if (!resp.ok) return;
        const data = await resp.json();
        const items = data.items || [];
        if (items.length && liveLog){
          const lines = items.map(it => {
            const base = JSON.stringify({ id: it.id, decision: it.decision, rationale: it.rationale });
            return it.retries ? `⚠ [retries:${it.retries}] ${base}` : base;
          });
          liveLog.textContent = (liveLog.textContent ? liveLog.textContent + "\n" : "") + lines.join("\n");
          if (liveLog.textContent.length > 100000) liveLog.textContent = liveLog.textContent.slice(-80000);
          liveLog.scrollTop = liveLog.scrollHeight;
        }
        delivered = data.next || delivered;
        if (data.status && (data.status === 'done' || data.status === 'error' || data.status === 'cancelled')){
          clearInterval(pollHandle); pollHandle = null;
        }
      } catch {}
    }, 1200);
  }
  function stopPolling(){ if (pollHandle) { clearInterval(pollHandle); pollHandle = null; } }
  // start polling by default in background; user can toggle visibility separately
  startPolling();
  const es = new EventSource(`/api/progress/${jobId}`);
  es.onmessage = (ev) => {
    try {
      const data = JSON.parse(ev.data);
      const processed = data.processed || 0; const total = data.total || 0; const pct = total ? Math.floor((processed/total)*100) : 0;
      bar.style.width = `${pct}%`;
      // Build label with live concurrency info
      let concLabel = "";
      if (data.concurrency) {
        const c = data.concurrency;
        const rlWarn = c.rate_limit_hits > 0 ? ` ⚠ ${c.rate_limit_hits} rate-limit(s)` : "";
        concLabel = ` | ⚡ ${c.current_concurrency} parallel${rlWarn}`;
      }
      label.textContent = `Processed ${processed} of ${total} (${pct}%)${concLabel}`;
      // SSE still updates the progress; detailed rows come via polling
      if (data.status === "done") { es.close(); label.textContent = `Completed: ${processed} of ${total} (100%)`; link.href = `/api/result/${jobId}`; link.classList.remove("hidden"); if(linkX){ linkX.href = `/api/result/${jobId}?format=xlsx`; linkX.classList.remove("hidden"); } if(btnCancel) btnCancel.disabled = true; if(btnRestart) btnRestart.classList.remove("hidden"); }
      if (data.status === "cancelled") { es.close(); label.textContent = `Cancelled at ${processed} of ${total}`; link.classList.add("hidden"); if(linkX) linkX.classList.add("hidden"); if(btnCancel) btnCancel.disabled = true; if(btnRestart) btnRestart.classList.remove("hidden"); }
      if (data.status === "error") { es.close(); showError("Backend reported an error (check server logs)."); if(btnCancel) btnCancel.disabled = true; }
    } catch {}
  };
  es.onerror = () => { showError("Lost connection to backend while streaming progress. Is the server running?"); };
  if (btnCancel) {
    btnCancel.disabled = false; btnCancel.onclick = async () => { try { btnCancel.disabled = true; label.textContent = `Cancelling…`; await fetch(`/api/cancel/${jobId}`, { method: "POST" }); } catch { btnCancel.disabled = false; } };
  }
  if (btnRestart){
    btnRestart.onclick = async () => {
      try {
        clearError();
        // Build a fresh payload so the user can change model/params
        // while reusing the same spreadsheet currently loaded in state
        const newPayload = buildPayload();
        btnRestart.disabled = true;
        label.textContent = "Restarting...";
        const resp = await fetch(`/api/start`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(newPayload) });
        if (!resp.ok) { const txt = await resp.text().catch(()=>""); showError(`Restart failed: ${resp.status} ${txt}`); btnRestart.disabled = false; return; }
        const data = await resp.json(); const newId = data.job_id; btnRestart.disabled = false; btnRestart.classList.add("hidden"); stopPolling(); showProgress(newId);
      } catch (e) { showError(e.message || String(e)); btnRestart.disabled = false; }
    };
  }
}

