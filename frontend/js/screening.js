// ── Screening tab — file upload, model params, criteria, payload, progress ────

// ── Error banner ──────────────────────────────────────────────────────────────

function showError(msg) {
  const b = document.getElementById("errorBanner");
  if (!b) { alert(msg); return; }
  b.textContent = msg;
  b.classList.remove("hidden");
}
function clearError() {
  const b = document.getElementById("errorBanner");
  if (b) { b.textContent = ""; b.classList.add("hidden"); }
}

// ── Model helpers ─────────────────────────────────────────────────────────────

function isReasoningModel(model) { return /^gpt-5/.test(model || ""); }
function isChatModel(model)      { return /^(gpt-4\.1|gpt-4o|gpt-3\.5)/.test(model || ""); }

function toggleParamGroups() {
  const model = document.getElementById("modelSelect")?.value || "";
  const r = document.getElementById("paramsReasoning");
  const c = document.getElementById("paramsChat");
  if (!r || !c) return;
  if (isReasoningModel(model))      { r.classList.remove("hidden"); c.classList.add("hidden"); }
  else if (isChatModel(model))      { c.classList.remove("hidden"); r.classList.add("hidden"); }
  else                              { r.classList.add("hidden");    c.classList.add("hidden"); }
}

function validateParamsOrWarn() {
  const model = document.getElementById("modelSelect")?.value || "";
  const warn = document.getElementById("paramsWarning");
  if (warn) { warn.classList.add("hidden"); warn.textContent = ""; }
  const num = id => {
    const el = document.getElementById(id); if (!el) return undefined;
    const v = el.value?.trim(); if (v === "" || v == null) return undefined;
    const n = Number(v); return Number.isFinite(n) ? n : NaN;
  };
  if (isChatModel(model)) {
    const t = num("temperature");
    if (t !== undefined && (t < 0 || t > 2)) {
      if (warn) { warn.textContent = "Temperature must be between 0 and 2."; warn.classList.remove("hidden"); }
      return false;
    }
  }
  return true;
}

function addParamTooltips() {
  const tips = {
    verbosity:        "Controls the level of detail in the response.\nlow: concise • medium: balanced (recommended) • high: very detailed.",
    reasoningEffort:  "Controls how much the model \"thinks\" before answering.\nminimal: fastest • low/medium: good balance (recommended) • high: most thorough but slower.",
    temperature:      "Controls randomness: 0 = deterministic, 2 = very creative. Recommended: 0.2–0.8.",
  };
  Object.entries(tips).forEach(([id, text]) => {
    const el = document.getElementById(id); if (!el) return;
    const label = el.closest("label.field"); if (!label) return;
    const titleSpan = label.querySelector("span"); if (!titleSpan) return;
    if (titleSpan.querySelector(".help")) return;
    const help = document.createElement("span"); help.className = "help"; help.tabIndex = 0;
    const icon = document.createElement("span"); icon.className = "icon"; icon.textContent = "?";
    const tip  = document.createElement("span"); tip.className  = "tip";  tip.textContent  = text;
    help.appendChild(icon); help.appendChild(tip); titleSpan.appendChild(help);
  });
}

// ── Provider selector ─────────────────────────────────────────────────────────

function initProviderSelector() {
  const btns = document.querySelectorAll(".provider-btn");
  const savedProvider = localStorage.getItem("provider") || "openai";
  state.provider = savedProvider;
  btns.forEach(btn => btn.classList.toggle("active", btn.dataset.provider === savedProvider));
  updateProviderUI();
  btns.forEach(btn => {
    btn.addEventListener("click", () => {
      state.provider = btn.dataset.provider;
      localStorage.setItem("provider", state.provider);
      btns.forEach(b => b.classList.toggle("active", b === btn));
      updateProviderUI();
    });
  });
}

function updateProviderUI() {
  const p = state.provider;
  ["openai", "anthropic", "google"].forEach(id => {
    const cfg = document.getElementById(`config-${id}`);
    if (cfg) cfg.classList.toggle("hidden", id !== p);
    const tier = document.getElementById(`tierSection-${id}`);
    if (tier) tier.classList.toggle("hidden", id !== p);
  });
}

// ── Tier presets & advanced settings ─────────────────────────────────────────

const TIER_PRESETS = {
  openai: {
    free: { concurrent: 2,  concurrent_max: 4,   concurrent_min: 1, record_retries: 3, aiup_after: 10, max_retries: 5, base_backoff: 2.0 },
    "1":  { concurrent: 5,  concurrent_max: 10,  concurrent_min: 1, record_retries: 3, aiup_after: 8,  max_retries: 5, base_backoff: 1.5 },
    "2":  { concurrent: 10, concurrent_max: 20,  concurrent_min: 2, record_retries: 3, aiup_after: 6,  max_retries: 5, base_backoff: 1.0 },
    "3":  { concurrent: 20, concurrent_max: 40,  concurrent_min: 2, record_retries: 3, aiup_after: 5,  max_retries: 5, base_backoff: 1.0 },
    "4":  { concurrent: 30, concurrent_max: 60,  concurrent_min: 3, record_retries: 3, aiup_after: 4,  max_retries: 5, base_backoff: 0.5 },
    "5":  { concurrent: 50, concurrent_max: 100, concurrent_min: 5, record_retries: 3, aiup_after: 3,  max_retries: 5, base_backoff: 0.5 },
  },
  anthropic: {
    starter:      { concurrent: 3,  concurrent_max: 5,  concurrent_min: 1, record_retries: 3, aiup_after: 8, max_retries: 5, base_backoff: 2.0 },
    standard:     { concurrent: 8,  concurrent_max: 15, concurrent_min: 1, record_retries: 3, aiup_after: 6, max_retries: 5, base_backoff: 1.5 },
    professional: { concurrent: 20, concurrent_max: 40, concurrent_min: 2, record_retries: 3, aiup_after: 5, max_retries: 5, base_backoff: 1.0 },
  },
  google: {
    free:        { concurrent: 2,  concurrent_max: 4,  concurrent_min: 1, record_retries: 3, aiup_after: 10, max_retries: 5, base_backoff: 2.0 },
    pay_per_use: { concurrent: 10, concurrent_max: 20, concurrent_min: 2, record_retries: 3, aiup_after: 6,  max_retries: 5, base_backoff: 1.0 },
    scale:       { concurrent: 30, concurrent_max: 60, concurrent_min: 3, record_retries: 3, aiup_after: 4,  max_retries: 5, base_backoff: 0.5 },
  },
};

function initAdvancedSettings() {
  const advIds = ["adv_concurrent","adv_concurrent_max","adv_concurrent_min",
                  "adv_record_retries","adv_aiup_after","adv_max_retries","adv_base_backoff"];
  advIds.forEach(id => {
    const v = localStorage.getItem(id); const el = document.getElementById(id);
    if (el && v != null) el.value = v;
  });

  const tierSel = document.getElementById("tierSelect");
  if (tierSel) {
    const savedTier = localStorage.getItem("tierSelect");
    if (savedTier) tierSel.value = savedTier;
    const hasAny = advIds.some(id => localStorage.getItem(id) != null);
    if (!hasAny) applyTierPreset("openai", tierSel.value || "3");
    tierSel.addEventListener("change", () => {
      localStorage.setItem("tierSelect", tierSel.value);
      applyTierPreset("openai", tierSel.value);
    });
  }

  const tierClaude = document.getElementById("tierSelectClaude");
  if (tierClaude) {
    const saved = localStorage.getItem("tierSelectClaude");
    if (saved) tierClaude.value = saved;
    tierClaude.addEventListener("change", () => {
      localStorage.setItem("tierSelectClaude", tierClaude.value);
      if (state.provider === "anthropic") applyTierPreset("anthropic", tierClaude.value);
    });
  }

  const tierGoogle = document.getElementById("tierSelectGoogle");
  if (tierGoogle) {
    const saved = localStorage.getItem("tierSelectGoogle");
    if (saved) tierGoogle.value = saved;
    tierGoogle.addEventListener("change", () => {
      localStorage.setItem("tierSelectGoogle", tierGoogle.value);
      if (state.provider === "google") applyTierPreset("google", tierGoogle.value);
    });
  }
}

function applyTierPreset(provider, tier) {
  const providerPresets = TIER_PRESETS[provider] || TIER_PRESETS.openai;
  const preset = providerPresets[tier]; if (!preset) return;
  const map = {
    adv_concurrent: "concurrent", adv_concurrent_max: "concurrent_max",
    adv_concurrent_min: "concurrent_min", adv_record_retries: "record_retries",
    adv_aiup_after: "aiup_after", adv_max_retries: "max_retries", adv_base_backoff: "base_backoff",
  };
  for (const [elId, key] of Object.entries(map)) {
    const el = document.getElementById(elId);
    if (el) { el.value = preset[key]; savePref(elId, preset[key]); }
  }
}

function getAdvancedParams() {
  const num = id => {
    const el = document.getElementById(id); if (!el) return undefined;
    const v = el.value?.trim(); if (v === "" || v == null) return undefined;
    const n = Number(v); return Number.isFinite(n) ? n : undefined;
  };
  const adv = {};
  const c    = num("adv_concurrent");      if (c    !== undefined) adv.concurrency        = c;
  const cm   = num("adv_concurrent_max");  if (cm   !== undefined) adv.concurrency_max    = cm;
  const cmin = num("adv_concurrent_min");  if (cmin !== undefined) adv.concurrent_min     = cmin;
  const rr   = num("adv_record_retries");  if (rr   !== undefined) adv.record_max_retries = rr;
  const aiup = num("adv_aiup_after");      if (aiup !== undefined) adv.aiup_after         = aiup;
  const mr   = num("adv_max_retries");     if (mr   !== undefined) adv.max_retries        = mr;
  const bb   = num("adv_base_backoff");    if (bb   !== undefined) adv.base_backoff       = bb;
  return adv;
}

// ── Criteria dynamic lists ────────────────────────────────────────────────────

function initCriteriaLists() {
  const incContainer = document.getElementById("inclusionList");
  const excContainer = document.getElementById("exclusionList");
  const addInc = document.getElementById("addInc");
  const addExc = document.getElementById("addExc");
  if (!incContainer || !excContainer) return;
  const savedInc = localStorage.getItem("inclusionCriteria") || "";
  const savedExc = localStorage.getItem("exclusionCriteria") || "";
  const incArr = savedInc.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const excArr = savedExc.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  if (!incArr.length) incArr.push("");
  if (!excArr.length) excArr.push("");
  incArr.forEach(v => addCriteriaInput(incContainer, v, "inc"));
  excArr.forEach(v => addCriteriaInput(excContainer, v, "exc"));
  if (addInc) addInc.onclick = () => addCriteriaInput(incContainer, "", "inc");
  if (addExc) addExc.onclick = () => addCriteriaInput(excContainer, "", "exc");
}

function addCriteriaInput(container, value, kind) {
  const row = document.createElement("div"); row.className = "criteria-row";
  const input = document.createElement("input"); input.type = "text";
  input.placeholder = kind === "inc"
    ? "e.g., population: rats or mice (preclinical)"
    : "e.g., case reports, reviews, editorials";
  input.value = value || "";
  input.addEventListener("input", persistCriteria);
  input.addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); addCriteriaInput(container, "", kind); }
  });
  const remove = document.createElement("button"); remove.type = "button";
  remove.className = "icon-btn remove-criterion"; remove.title = "Remove this criterion";
  remove.setAttribute("aria-label", "Remove"); remove.textContent = "×";
  remove.addEventListener("click", () => {
    const total = container.querySelectorAll("input").length;
    if (total > 1) container.removeChild(row); else input.value = "";
    persistCriteria();
  });
  row.appendChild(input); row.appendChild(remove); container.appendChild(row);
  setTimeout(() => input.focus(), 0);
}

function persistCriteria() {
  const incVals = Array.from(document.querySelectorAll("#inclusionList input")).map(el => el.value.trim()).filter(Boolean);
  const excVals = Array.from(document.querySelectorAll("#exclusionList input")).map(el => el.value.trim()).filter(Boolean);
  localStorage.setItem("inclusionCriteria", incVals.join("\n"));
  localStorage.setItem("exclusionCriteria", excVals.join("\n"));
}

// ── File upload & parsing ─────────────────────────────────────────────────────

let lastPayload = null;

const fileInput       = document.getElementById("fileInput");
const dropzone        = document.getElementById("dropzone");
const fileMeta        = document.getElementById("fileMeta");
const sheetPickerWrap = document.getElementById("sheetPickerWrap");
const sheetSelect     = document.getElementById("sheetSelect");
const columnsStatus   = document.getElementById("columnsStatus");
const previewWrap     = document.getElementById("previewWrap");
const previewTable    = document.getElementById("previewTable");
const errorBanner     = document.getElementById("errorBanner");

dropzone.addEventListener("dragover",  e => { e.preventDefault(); dropzone.classList.add("hover"); });
dropzone.addEventListener("dragleave", ()  => dropzone.classList.remove("hover"));
dropzone.addEventListener("drop", e => {
  e.preventDefault(); dropzone.classList.remove("hover");
  const f = e.dataTransfer.files?.[0]; if (f) handleFile(f);
});
fileInput.addEventListener("change", e => { const f = e.target.files?.[0]; if (f) handleFile(f); });

async function ensureXLSX() {
  if (typeof XLSX !== "undefined") return true;
  const trySrc = src => new Promise((resolve, reject) => {
    const s = document.createElement("script"); s.src = src; s.defer = true;
    s.onload = resolve; s.onerror = reject; document.head.appendChild(s);
  });
  try { await trySrc("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"); if (typeof XLSX !== "undefined") return true; } catch {}
  try { await trySrc("https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js"); if (typeof XLSX !== "undefined") return true; } catch {}
  return false;
}

async function handleFile(file) {
  state.filename = file.name;
  fileMeta.textContent = `${file.name} • ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  const outEl = document.getElementById("outputFilename");
  if (outEl) outEl.value = file.name.replace(/\.[^.]+$/, "");
  const reader = new FileReader();
  reader.onload = async evt => {
    try {
      const data = evt.target.result;
      const lower = file.name.toLowerCase();
      if (lower.endsWith(".csv")) {
        const text = decodeTextWithFallback(data);
        const rows = csvToArray(text);
        if (!rows.length) showError("No rows could be parsed from the CSV. Check delimiter, encoding and header row.");
        state.workbook = null; state.rows = rows;
        sheetPickerWrap.classList.add("hidden"); validateAndPreview();
      } else if (lower.endsWith(".xlsx") || lower.endsWith(".xls")) {
        if (typeof XLSX === "undefined") {
          const ok = await ensureXLSX();
          if (!ok) { showError("Excel library not loaded. Ensure internet or load the page via http://localhost:8000."); return; }
        }
        const wb = XLSX.read(data, { type: "array" });
        if (!wb.SheetNames?.length) { showError("No sheet found in this Excel file."); return; }
        state.workbook = wb; sheetSelect.innerHTML = "";
        wb.SheetNames.forEach((s, i) => {
          const opt = document.createElement("option"); opt.value = s; opt.textContent = s;
          if (i === 0) opt.selected = true; sheetSelect.appendChild(opt);
        });
        sheetPickerWrap.classList.remove("hidden");
        state.selectedSheet = wb.SheetNames[0]; loadWorkbookSheet(state.selectedSheet);
      } else {
        showError("Unsupported file type. Upload .csv, .xlsx, or .xls.");
      }
    } catch (e) { showError(`Failed to read file: ${e?.message || e}`); }
  };
  reader.readAsArrayBuffer(file);
}

sheetSelect.addEventListener("change", () => { state.selectedSheet = sheetSelect.value; loadWorkbookSheet(state.selectedSheet); });

function loadWorkbookSheet(name) {
  const ws = state.workbook.Sheets[name];
  state.rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  validateAndPreview();
}

// ── Validation & preview ──────────────────────────────────────────────────────

function validateAndPreview() {
  const { cols, rows: normalized } = inferColumns(state.rows);
  state.rows = normalized;

  const hasTitle    = cols.includes("title");
  const hasAbstract = cols.includes("abstract");
  let statusHtml = "";
  if (hasTitle && hasAbstract) {
    statusHtml = `<strong style="color:var(--ok)">✓</strong> Detected required columns: <code>title</code>, <code>abstract</code>.`;
  } else if (!state.rows.length) {
    statusHtml = "No rows could be read from this sheet/file.";
  } else {
    const miss = []; if (!hasTitle) miss.push("title"); if (!hasAbstract) miss.push("abstract");
    statusHtml = `<span class="badge bad">Missing</span> Required columns not found: <code>${miss.join("</code>, <code>")}</code>. Expected: title/título/titulo and abstract/resumo/summary.`;
  }
  columnsStatus.innerHTML = statusHtml;

  const head = previewTable.querySelector("thead"); const body = previewTable.querySelector("tbody");
  head.innerHTML = ""; body.innerHTML = "";
  if (!state.rows.length) { previewWrap.classList.add("hidden"); return; }

  const colsToShow = cols.length ? cols : Object.keys(state.rows[0] ?? {});
  const trh = document.createElement("tr");
  colsToShow.forEach(c => { const th = document.createElement("th"); th.textContent = c; trh.appendChild(th); });
  head.appendChild(trh);

  const ell = (s, n = 200) => { if (s == null) return ""; const str = String(s); return str.length > n ? str.slice(0, n - 1) + "…" : str; };
  state.rows.slice(0, 5).forEach(r => {
    const tr = document.createElement("tr");
    colsToShow.forEach(c => {
      const td = document.createElement("td"); const full = r[c] ?? "";
      td.textContent = ell(full); if (String(full).length > 200) td.title = String(full);
      tr.appendChild(td);
    });
    body.appendChild(tr);
  });
  previewWrap.classList.remove("hidden");
  document.getElementById("btnSend").disabled = !(hasTitle && hasAbstract);
  if (state.filename) {
    const sizePart = (fileMeta.textContent.match(/•\s[\d.]+\sMB/) || [""])[0];
    fileMeta.textContent = sizePart
      ? `${state.filename} ${sizePart} • ${state.rows.length.toLocaleString()} rows`
      : `${state.filename} • ${state.rows.length.toLocaleString()} rows`;
  }
}

// ── Payload builder ───────────────────────────────────────────────────────────

function buildPayload() {
  const provider = state.provider || "openai";
  const synopsis = document.getElementById("studySynopsis")?.value || "";
  const incListEls = document.querySelectorAll("#inclusionList input");
  const excListEls = document.querySelectorAll("#exclusionList input");
  const inclusionArr = incListEls.length
    ? Array.from(incListEls).map(el => el.value.trim()).filter(Boolean)
    : splitLines(document.getElementById("inclusionCriteria")?.value || "");
  const exclusionArr = excListEls.length
    ? Array.from(excListEls).map(el => el.value.trim()).filter(Boolean)
    : splitLines(document.getElementById("exclusionCriteria")?.value || "");

  let model, api_key;
  const params = {};

  if (provider === "anthropic") {
    model   = document.getElementById("modelSelectClaude")?.value || "";
    api_key = (document.getElementById("apiKeyClaude")?.value || "").trim();
    if (!model)   throw new Error("Select a Claude model.");
    if (!api_key) throw new Error("Enter your Anthropic API key.");
    const t = document.getElementById("temperatureClaude")?.value?.trim();
    if (t) params.temperature = Number(t);
  } else if (provider === "google") {
    model   = document.getElementById("modelSelectGoogle")?.value || "";
    api_key = (document.getElementById("apiKeyGoogle")?.value || "").trim();
    if (!model)   throw new Error("Select a Gemini model.");
    if (!api_key) throw new Error("Enter your Google API key.");
    const t = document.getElementById("temperatureGoogle")?.value?.trim();
    if (t) params.temperature = Number(t);
  } else {
    model   = document.getElementById("modelSelect")?.value || "";
    api_key = (document.getElementById("apiKey")?.value || "").trim();
    if (!model)   throw new Error("Select a model.");
    if (!api_key) throw new Error("Enter your OpenAI API key.");
    if (!validateParamsOrWarn()) throw new Error("Please fix parameters.");
    if (isReasoningModel(model)) {
      const v  = document.getElementById("verbosity")?.value?.trim();
      const re = document.getElementById("reasoningEffort")?.value?.trim();
      if (v)  params.verbosity        = v;
      if (re) params.reasoning_effort = re;
    } else if (isChatModel(model)) {
      const t = document.getElementById("temperature")?.value?.trim();
      if (t) params.temperature = Number(t);
    }
  }

  if (!state.rows.length) throw new Error("Load a spreadsheet.");
  const first = state.rows[0] ?? {};
  if (!("title" in first) || !("abstract" in first))
    throw new Error("The spreadsheet must have 'title' and 'abstract' columns (or common variants).");

  Object.assign(params, getAdvancedParams());

  return {
    provider, model, api_key,
    study_synopsis: synopsis,
    inclusion_criteria: inclusionArr,
    exclusion_criteria: exclusionArr,
    params,
    sheet:    state.selectedSheet || state.filename || "",
    filename: state.filename || "",
    sample_preview: state.rows.slice(0, 5).map(r => ({ title: r.title ?? "", abstract: r.abstract ?? "" })),
    normalized_columns: true,
    records: state.rows.map((r, i) => ({ id: r.id ?? r.ID ?? r.Id ?? i + 1, title: r.title ?? "", abstract: r.abstract ?? "" })),
  };
}

// ── Action buttons ────────────────────────────────────────────────────────────

document.getElementById("btnGenerate").addEventListener("click", () => {
  try {
    const payload = buildPayload();
    document.getElementById("payloadOut").textContent = JSON.stringify(payload, null, 2);
    document.getElementById("payloadOut").classList.remove("hidden");
    const btnCP = document.getElementById("btnCollapsePayload");
    if (btnCP) btnCP.textContent = "Collapse";
    document.getElementById("payloadCard").classList.remove("hidden");
  } catch (e) { showError(e.message); }
});

document.getElementById("btnCollapsePayload").addEventListener("click", () => {
  const pre = document.getElementById("payloadOut");
  const btn = document.getElementById("btnCollapsePayload");
  const collapsed = pre.classList.toggle("hidden");
  btn.textContent = collapsed ? "Expand" : "Collapse";
});

document.getElementById("btnCopy").addEventListener("click", async () => {
  try {
    await navigator.clipboard.writeText(document.getElementById("payloadOut").textContent);
    const btn = document.getElementById("btnCopy"); btn.textContent = "Copied!";
    setTimeout(() => btn.textContent = "Copy", 1200);
  } catch {
    const ta = document.createElement("textarea"); ta.value = document.getElementById("payloadOut").textContent;
    document.body.appendChild(ta); ta.select(); document.execCommand("copy"); document.body.removeChild(ta);
  }
});

document.getElementById("btnSend").addEventListener("click", async () => {
  try {
    clearError();
    const payload = buildPayload(); lastPayload = payload;
    const btn = document.getElementById("btnSend");
    const prevHTML = btn.innerHTML; btn.textContent = "Sending…"; btn.disabled = true;
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 30000);
    const resp = await fetch("/api/start", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload), signal: controller.signal,
    }).finally(() => clearTimeout(timeout));
    if (!resp.ok) {
      const txt = await resp.text().catch(() => "");
      let friendly = `Send failed (${resp.status}).`;
      if (resp.status === 501) friendly += " Run 'uvicorn backend.server:app --port 8000' and open http://localhost:8000.";
      else if (resp.status === 404) friendly += " Endpoint not found. Access via http://localhost:8000.";
      else if (resp.status === 405) friendly += " Method not allowed — likely hitting a static server, not FastAPI.";
      showError(friendly + (txt ? `\nDetails: ${txt}` : ""));
      btn.innerHTML = prevHTML; btn.disabled = false;
      return;
    }
    const data = await resp.json(); showProgress(data.job_id);
    btn.innerHTML = prevHTML; btn.disabled = false;
  } catch (e) { showError(e.message || String(e)); document.getElementById("btnSend").disabled = false; }
});

document.getElementById("btnClear").addEventListener("click", () => {
  state.workbook = null; state.filename = null; state.selectedSheet = null; state.rows = [];
  const fi = document.getElementById("fileInput"); if (fi) fi.value = "";
  fileMeta.textContent = "";
  const outEl2 = document.getElementById("outputFilename"); if (outEl2) outEl2.value = "";
  sheetPickerWrap.classList.add("hidden");
  columnsStatus.textContent = "Waiting for file...";
  previewTable.querySelector("thead").innerHTML = "";
  previewTable.querySelector("tbody").innerHTML = "";
  previewWrap.classList.add("hidden");
  document.getElementById("payloadCard").classList.add("hidden");
  document.getElementById("btnSend").disabled = true;
  lastPayload = null;
});

// ── Progress (SSE + polling) ──────────────────────────────────────────────────

function showProgress(jobId) {
  const card        = document.getElementById("progressCard");
  const bar         = document.getElementById("progressBar");
  const label       = document.getElementById("progressLabel");
  const link        = document.getElementById("downloadLink");
  const linkX       = document.getElementById("downloadLinkXlsx");
  const btnCancel   = document.getElementById("btnCancel");
  const btnRestart  = document.getElementById("btnRestart");
  const liveLog     = document.getElementById("liveLog");
  const btnToggleLog= document.getElementById("btnToggleLog");
  const elapsedEl   = document.getElementById("elapsedLabel");

  let elapsedInterval = null;
  let elapsedStart = null;

  function startElapsed() {
    if (elapsedInterval || !elapsedEl) return;
    elapsedStart = Date.now();
    elapsedEl.classList.remove("hidden");
    elapsedEl.textContent = "Time elapsed: 0s";
    elapsedInterval = setInterval(() => {
      const s = Math.floor((Date.now() - elapsedStart) / 1000);
      const m = Math.floor(s / 60);
      elapsedEl.textContent = m > 0 ? `Time elapsed: ${m}m ${s % 60}s` : `Time elapsed: ${s}s`;
    }, 1000);
  }
  function stopElapsed() {
    if (elapsedInterval) { clearInterval(elapsedInterval); elapsedInterval = null; }
    if (elapsedEl && elapsedStart) {
      const s = Math.floor((Date.now() - elapsedStart) / 1000);
      const m = Math.floor(s / 60);
      elapsedEl.textContent = `Total analysis time: ${m > 0 ? `${m}m ${s % 60}s` : `${s}s`}`;
    }
  }

  card.classList.remove("hidden"); link.classList.add("hidden");
  if (linkX) linkX.classList.add("hidden");
  bar.style.width = "0%"; label.textContent = "Starting...";
  if (elapsedEl) { elapsedEl.textContent = ""; elapsedEl.classList.add("hidden"); }
  if (liveLog) { liveLog.textContent = ""; liveLog.classList.add("hidden"); }
  if (btnToggleLog) {
    btnToggleLog.textContent = "Show live log";
    btnToggleLog.onclick = () => {
      if (!liveLog) return;
      const hidden = liveLog.classList.contains("hidden");
      liveLog.classList.toggle("hidden", !hidden);
      btnToggleLog.textContent = hidden ? "Hide live log" : "Show live log";
    };
  }

  let pollHandle = null; let delivered = 0;
  function startPolling() {
    if (pollHandle) return;
    pollHandle = setInterval(async () => {
      try {
        const resp = await fetch(`/api/partial/${jobId}?since=${delivered}`);
        if (!resp.ok) return;
        const data = await resp.json();
        const items = data.items || [];
        if (items.length && liveLog) {
          const lines = items.map(it => {
            const base = JSON.stringify({ id: it.id, decision: it.decision, rationale: it.rationale });
            return it.retries ? `⚠ [retries:${it.retries}] ${base}` : base;
          });
          liveLog.textContent = (liveLog.textContent ? liveLog.textContent + "\n" : "") + lines.join("\n");
          if (liveLog.textContent.length > 100000) liveLog.textContent = liveLog.textContent.slice(-80000);
          liveLog.scrollTop = liveLog.scrollHeight;
        }
        delivered = data.next || delivered;
        if (data.status && ["done", "error", "cancelled"].includes(data.status)) {
          clearInterval(pollHandle); pollHandle = null;
        }
      } catch {}
    }, 1200);
  }
  function stopPolling() { if (pollHandle) { clearInterval(pollHandle); pollHandle = null; } }
  startPolling();

  const es = new EventSource(`/api/progress/${jobId}`);
  es.onmessage = ev => {
    try {
      const data = JSON.parse(ev.data);
      const processed = data.processed || 0; const total = data.total || 0;
      if (!elapsedStart && processed > 0) startElapsed();
      const pct = total ? Math.floor((processed / total) * 100) : 0;
      bar.style.width = `${pct}%`;
      let concLabel = "";
      if (data.concurrency) {
        const c = data.concurrency;
        const rlWarn = c.rate_limit_hits > 0 ? ` ⚠ ${c.rate_limit_hits} rate-limit(s)` : "";
        concLabel = ` | ⚡ ${c.current_concurrency} parallel${rlWarn}`;
      }
      label.textContent = `Processed ${processed} of ${total} (${pct}%)${concLabel}`;
      if (data.status === "done") {
        es.close(); stopPolling(); stopElapsed();
        label.textContent = `Completed: ${processed} of ${total} (100%)`;
        const outEl = document.getElementById("outputFilename");
        const baseName = (outEl?.value?.trim() || "screening_result").replace(/[/\\?%*:|"<>]/g, "_");
        link.href = `/api/result/${jobId}`; link.download = `${baseName}.csv`; link.classList.remove("hidden");
        if (linkX) { linkX.href = `/api/result/${jobId}?format=xlsx`; linkX.download = `${baseName}.xlsx`; linkX.classList.remove("hidden"); }
        if (btnCancel) btnCancel.disabled = true;
        if (btnRestart) btnRestart.classList.remove("hidden");
      }
      if (data.status === "cancelled") {
        es.close(); stopPolling(); stopElapsed();
        label.textContent = `Cancelled at ${processed} of ${total}`;
        link.classList.add("hidden"); if (linkX) linkX.classList.add("hidden");
        if (btnCancel) btnCancel.disabled = true;
        if (btnRestart) btnRestart.classList.remove("hidden");
      }
      if (data.status === "error") { es.close(); stopPolling(); stopElapsed(); showError("Backend reported an error (check server logs)."); if (btnCancel) btnCancel.disabled = true; }
    } catch {}
  };
  es.onerror = () => showError("Lost connection to backend while streaming progress. Is the server running?");

  if (btnCancel) {
    btnCancel.disabled = false;
    btnCancel.onclick = async () => {
      try { btnCancel.disabled = true; label.textContent = "Cancelling…"; await fetch(`/api/cancel/${jobId}`, { method: "POST" }); }
      catch { btnCancel.disabled = false; }
    };
  }
  if (btnRestart) {
    btnRestart.onclick = async () => {
      try {
        clearError();
        const newPayload = buildPayload(); btnRestart.disabled = true; label.textContent = "Restarting...";
        const resp = await fetch("/api/start", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(newPayload) });
        if (!resp.ok) { const txt = await resp.text().catch(() => ""); showError(`Restart failed: ${resp.status} ${txt}`); btnRestart.disabled = false; return; }
        const data = await resp.json(); btnRestart.disabled = false; btnRestart.classList.add("hidden");
        stopPolling(); showProgress(data.job_id);
      } catch (e) { showError(e.message || String(e)); btnRestart.disabled = false; }
    };
  }
}
