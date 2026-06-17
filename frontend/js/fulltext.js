// ── Full-text screening tab — PDF upload, prompt preview, send & progress ────
//
// Independent from the TIAB screening tab: own provider/model/criteria stored
// under "ft_*" localStorage keys. PDFs are uploaded via multipart/form-data
// to /api/fulltext/start and the LLM reads the PDF natively (no extraction).

const ftState = {
  provider: "openai",
  files: [],          // [{file: File, name: string, size: number}]
  jobId: null,
  sse: null,
  startedAt: 0,
};

// ── Helpers ───────────────────────────────────────────────────────────────────

function ftShowError(msg) {
  const b = document.getElementById("ftErrorBanner");
  if (!b) { alert(msg); return; }
  b.textContent = msg;
  b.classList.remove("hidden");
}
function ftClearError() {
  const b = document.getElementById("ftErrorBanner");
  if (b) { b.textContent = ""; b.classList.add("hidden"); }
}
function ftIsReasoningModel(model) { return /^gpt-5/.test(model || ""); }
function ftIsChatModel(model)      { return /^(gpt-4\.1|gpt-4o|gpt-3\.5)/.test(model || ""); }

function ftHumanSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

// ── Provider selector ─────────────────────────────────────────────────────────

function ftInitProviderSelector() {
  const sel = document.getElementById("ftProviderSelector");
  if (!sel) return;
  const btns = sel.querySelectorAll(".provider-btn");
  const saved = localStorage.getItem("ft_provider") || "openai";
  ftState.provider = saved;
  btns.forEach(btn => btn.classList.toggle("active", btn.dataset.provider === saved));
  ftUpdateProviderUI();
  btns.forEach(btn => {
    btn.addEventListener("click", () => {
      ftState.provider = btn.dataset.provider;
      localStorage.setItem("ft_provider", ftState.provider);
      btns.forEach(b => b.classList.toggle("active", b === btn));
      ftUpdateProviderUI();
    });
  });
}

function ftUpdateProviderUI() {
  const p = ftState.provider;
  ["openai", "anthropic", "google"].forEach(id => {
    const cfg = document.getElementById(`ft-config-${id}`);
    if (cfg) cfg.classList.toggle("hidden", id !== p);
  });
  ftToggleParamGroups();
}

function ftToggleParamGroups() {
  const model = document.getElementById("ftModelSelect")?.value || "";
  const r = document.getElementById("ftParamsReasoning");
  const c = document.getElementById("ftParamsChat");
  if (!r || !c) return;
  if (ftIsReasoningModel(model))      { r.classList.remove("hidden"); c.classList.add("hidden"); }
  else if (ftIsChatModel(model))      { c.classList.remove("hidden"); r.classList.add("hidden"); }
  else                                { r.classList.add("hidden");    c.classList.add("hidden"); }
}

// ── Criteria lists (independent from TIAB) ───────────────────────────────────

function ftInitCriteriaLists() {
  const incContainer = document.getElementById("ftInclusionList");
  const excContainer = document.getElementById("ftExclusionList");
  const addInc = document.getElementById("ftAddInc");
  const addExc = document.getElementById("ftAddExc");
  if (!incContainer || !excContainer) return;
  const savedInc = localStorage.getItem("ft_inclusionCriteria") || "";
  const savedExc = localStorage.getItem("ft_exclusionCriteria") || "";
  const incArr = savedInc.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const excArr = savedExc.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  if (!incArr.length) incArr.push("");
  if (!excArr.length) excArr.push("");
  incArr.forEach(v => ftAddCriteriaInput(incContainer, v, "inc"));
  excArr.forEach(v => ftAddCriteriaInput(excContainer, v, "exc"));
  if (addInc) addInc.onclick = () => ftAddCriteriaInput(incContainer, "", "inc");
  if (addExc) addExc.onclick = () => ftAddCriteriaInput(excContainer, "", "exc");
}

function ftAddCriteriaInput(container, value, kind) {
  const row = document.createElement("div"); row.className = "criteria-row";
  const input = document.createElement("input"); input.type = "text";
  input.placeholder = kind === "inc"
    ? "e.g., RCT design with at least 12-week follow-up"
    : "e.g., abstract-only / conference proceedings without full data";
  input.value = value || "";
  input.addEventListener("input", ftPersistCriteria);
  input.addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); ftAddCriteriaInput(container, "", kind); }
  });
  const remove = document.createElement("button"); remove.type = "button";
  remove.className = "icon-btn remove-criterion"; remove.title = "Remove this criterion";
  remove.setAttribute("aria-label", "Remove"); remove.textContent = "×";
  remove.addEventListener("click", () => {
    const total = container.querySelectorAll("input").length;
    if (total > 1) container.removeChild(row); else input.value = "";
    ftPersistCriteria();
  });
  row.appendChild(input); row.appendChild(remove); container.appendChild(row);
  setTimeout(() => input.focus(), 0);
}

function ftPersistCriteria() {
  const incVals = Array.from(document.querySelectorAll("#ftInclusionList input")).map(el => el.value.trim()).filter(Boolean);
  const excVals = Array.from(document.querySelectorAll("#ftExclusionList input")).map(el => el.value.trim()).filter(Boolean);
  localStorage.setItem("ft_inclusionCriteria", incVals.join("\n"));
  localStorage.setItem("ft_exclusionCriteria", excVals.join("\n"));
}

function ftReadCriteria() {
  const inc = Array.from(document.querySelectorAll("#ftInclusionList input")).map(el => el.value.trim()).filter(Boolean);
  const exc = Array.from(document.querySelectorAll("#ftExclusionList input")).map(el => el.value.trim()).filter(Boolean);
  const synopsis = document.getElementById("ftStudySynopsis")?.value || "";
  return { synopsis, inc, exc };
}

// ── File upload ──────────────────────────────────────────────────────────────

function ftInitUpload() {
  const dropzone = document.getElementById("ftDropzone");
  const fileInput = document.getElementById("ftFileInput");
  if (!dropzone || !fileInput) return;

  dropzone.addEventListener("dragover",  e => { e.preventDefault(); dropzone.classList.add("hover"); });
  dropzone.addEventListener("dragleave", ()  => dropzone.classList.remove("hover"));
  dropzone.addEventListener("drop", e => {
    e.preventDefault(); dropzone.classList.remove("hover");
    if (e.dataTransfer.files?.length) ftAddFiles(e.dataTransfer.files);
  });
  fileInput.addEventListener("change", e => {
    if (e.target.files?.length) ftAddFiles(e.target.files);
    fileInput.value = "";
  });

  const clearBtn = document.getElementById("ftBtnClear");
  if (clearBtn) clearBtn.onclick = ftClearFiles;
}

function ftAddFiles(fileList) {
  const MAX_BYTES = 32 * 1024 * 1024;
  let rejected = 0;
  for (const f of fileList) {
    if (!f.name.toLowerCase().endsWith(".pdf")) { rejected++; continue; }
    if (f.size > MAX_BYTES) { rejected++; continue; }
    if (ftState.files.some(x => x.name === f.name && x.size === f.size)) continue;
    ftState.files.push({ file: f, name: f.name, size: f.size });
  }
  if (rejected > 0) {
    ftShowError(`${rejected} file(s) ignored — must be PDF and ≤ 32 MB.`);
  } else {
    ftClearError();
  }
  ftRenderFileList();
}

function ftClearFiles() {
  ftState.files = [];
  ftRenderFileList();
  ftClearError();
}

function ftRenderFileList() {
  const list = document.getElementById("ftFileList");
  if (!list) return;
  list.innerHTML = "";
  if (ftState.files.length === 0) {
    list.classList.add("hidden");
  } else {
    list.classList.remove("hidden");
    ftState.files.forEach((entry, idx) => {
      const row = document.createElement("div");
      row.className = "ft-file-item";
      row.innerHTML = `
        <span class="ft-file-name" title="${entry.name}">${entry.name}</span>
        <span class="ft-file-size">${ftHumanSize(entry.size)}</span>
      `;
      const btn = document.createElement("button");
      btn.type = "button"; btn.className = "ft-file-remove"; btn.title = "Remove";
      btn.textContent = "×";
      btn.addEventListener("click", () => {
        ftState.files.splice(idx, 1);
        ftRenderFileList();
      });
      row.appendChild(btn);
      list.appendChild(row);
    });
  }
  const sendBtn = document.getElementById("ftBtnSend");
  if (sendBtn) sendBtn.disabled = ftState.files.length === 0;
}

// ── Prompt preview modal ─────────────────────────────────────────────────────

function ftInitPromptPreview() {
  const btn = document.getElementById("ftPreviewPromptBtn");
  if (btn) btn.addEventListener("click", ftShowPromptPreview);

  const modal = document.getElementById("ftPromptModal");
  if (modal) {
    modal.querySelectorAll("[data-close-modal]").forEach(el => {
      el.addEventListener("click", () => modal.classList.add("hidden"));
    });
  }
  const copyBtn = document.getElementById("ftCopyPromptBtn");
  if (copyBtn) {
    copyBtn.addEventListener("click", async () => {
      const txt = document.getElementById("ftPromptPreviewText")?.textContent || "";
      try {
        await navigator.clipboard.writeText(txt);
        copyBtn.textContent = "Copied!";
        setTimeout(() => { copyBtn.textContent = "Copy"; }, 1500);
      } catch {
        copyBtn.textContent = "Copy failed";
      }
    });
  }
}

async function ftShowPromptPreview() {
  ftClearError();
  const { synopsis, inc, exc } = ftReadCriteria();
  const firstName = ftState.files[0]?.name || "<article.pdf>";
  try {
    const r = await fetch("/api/fulltext/preview-prompt", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        study_synopsis: synopsis,
        inclusion_criteria: inc,
        exclusion_criteria: exc,
        filename: firstName,
      }),
    });
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const data = await r.json();
    const txt = data.prompt || "(empty)";
    const out = document.getElementById("ftPromptPreviewText");
    if (out) out.textContent = txt;
    const modal = document.getElementById("ftPromptModal");
    if (modal) modal.classList.remove("hidden");
  } catch (e) {
    ftShowError(`Could not load prompt preview: ${e.message || e}`);
  }
}

// ── Send job ─────────────────────────────────────────────────────────────────

function ftCurrentApiKey() {
  switch (ftState.provider) {
    case "anthropic": return document.getElementById("ftApiKeyClaude")?.value || "";
    case "google":    return document.getElementById("ftApiKeyGoogle")?.value || "";
    default:          return document.getElementById("ftApiKey")?.value || "";
  }
}
function ftCurrentModel() {
  switch (ftState.provider) {
    case "anthropic": return document.getElementById("ftModelSelectClaude")?.value || "";
    case "google":    return document.getElementById("ftModelSelectGoogle")?.value || "";
    default:          return document.getElementById("ftModelSelect")?.value || "";
  }
}
function ftCurrentParams() {
  const params = {};
  if (ftState.provider === "openai") {
    const model = ftCurrentModel();
    if (ftIsReasoningModel(model)) {
      const v = document.getElementById("ftVerbosity")?.value;
      const e = document.getElementById("ftReasoningEffort")?.value;
      if (v) params.verbosity = v;
      if (e) params.reasoning_effort = e;
    } else if (ftIsChatModel(model)) {
      const t = document.getElementById("ftTemperature")?.value;
      if (t !== "") params.temperature = parseFloat(t);
    }
  } else if (ftState.provider === "anthropic") {
    const t = document.getElementById("ftTemperatureClaude")?.value;
    if (t !== "") params.temperature = parseFloat(t);
  } else if (ftState.provider === "google") {
    const t = document.getElementById("ftTemperatureGoogle")?.value;
    if (t !== "") params.temperature = parseFloat(t);
  }
  // Beta defaults — conservative concurrency since PDF calls are heavy
  params.concurrency = 3;
  params.concurrency_max = 6;
  params.concurrent_min = 1;
  return params;
}

async function ftSendJob() {
  ftClearError();
  if (ftState.files.length === 0) { ftShowError("Add at least one PDF first."); return; }
  const apiKey = ftCurrentApiKey();
  if (!apiKey) { ftShowError("API key is missing — paste it in the configuration card."); return; }
  const { synopsis, inc, exc } = ftReadCriteria();
  if (!synopsis) { ftShowError("Add a Synopsis / PICO before sending."); return; }

  const fd = new FormData();
  fd.append("provider", ftState.provider);
  fd.append("model", ftCurrentModel());
  fd.append("api_key", apiKey);
  fd.append("study_synopsis", synopsis);
  fd.append("inclusion_criteria", JSON.stringify(inc));
  fd.append("exclusion_criteria", JSON.stringify(exc));
  fd.append("params", JSON.stringify(ftCurrentParams()));
  ftState.files.forEach(entry => fd.append("pdfs", entry.file, entry.name));

  ftShowProgressUI();
  try {
    const r = await fetch("/api/fulltext/start", { method: "POST", body: fd });
    if (!r.ok) {
      const t = await r.text();
      throw new Error(`HTTP ${r.status}: ${t}`);
    }
    const data = await r.json();
    ftState.jobId = data.job_id;
    ftState.startedAt = Date.now();
    ftStartProgressStream(data.job_id);
  } catch (e) {
    ftHideProgressUI();
    ftShowError(`Could not start full-text job: ${e.message || e}`);
  }
}

// ── Progress UI ──────────────────────────────────────────────────────────────

function ftShowProgressUI() {
  document.getElementById("ftProgressCard")?.classList.remove("hidden");
  document.getElementById("ftDownloadLink")?.classList.add("hidden");
  document.getElementById("ftDownloadCsvLink")?.classList.add("hidden");
  document.getElementById("ftBtnRestart")?.classList.add("hidden");
  document.getElementById("ftBtnCancel").disabled = false;
  document.getElementById("ftLiveLog").textContent = "";
  document.getElementById("ftProgressBar").style.width = "0%";
  document.getElementById("ftProgressLabel").textContent = "Starting…";
  document.getElementById("ftBtnSend").disabled = true;
}

function ftHideProgressUI() {
  document.getElementById("ftBtnSend").disabled = ftState.files.length === 0;
}

function ftStartProgressStream(jobId) {
  if (ftState.sse) { try { ftState.sse.close(); } catch {} }
  const es = new EventSource(`/api/fulltext/progress/${jobId}`);
  ftState.sse = es;
  es.onmessage = ev => {
    let msg; try { msg = JSON.parse(ev.data); } catch { return; }
    const total = msg.total || 0;
    const proc = msg.processed || 0;
    const pct = total > 0 ? Math.min(100, Math.round(proc / total * 100)) : 0;
    const bar = document.getElementById("ftProgressBar");
    if (bar) bar.style.width = `${pct}%`;
    const lbl = document.getElementById("ftProgressLabel");
    if (lbl) {
      const elapsed = Math.floor((Date.now() - ftState.startedAt) / 1000);
      lbl.textContent = `${proc}/${total}  (${pct}%) — ${elapsed}s elapsed`;
    }
    if (msg.last) ftAppendLogLine(msg.last);
    if (msg.error) ftShowError(msg.error);
    if (msg.status === "done" || msg.status === "cancelled" || msg.status === "error") {
      try { es.close(); } catch {}
      ftState.sse = null;
      ftOnJobFinished(msg.status);
    }
  };
  es.onerror = () => {
    try { es.close(); } catch {}
    ftState.sse = null;
  };
}

function ftAppendLogLine(last) {
  const log = document.getElementById("ftLiveLog");
  if (!log) return;
  const line = `[${last.id}] ${last.decision?.toUpperCase() || "?"} — ${last.filename || ""}: ${last.rationale || ""}`;
  log.textContent += line + "\n";
  log.scrollTop = log.scrollHeight;
}

function ftOnJobFinished(status) {
  document.getElementById("ftBtnCancel").disabled = true;
  document.getElementById("ftBtnSend").disabled = false;
  document.getElementById("ftBtnRestart").classList.remove("hidden");
  if (status === "done" && ftState.jobId) {
    const xlsx = document.getElementById("ftDownloadLink");
    const csv = document.getElementById("ftDownloadCsvLink");
    if (xlsx) { xlsx.href = `/api/fulltext/result/${ftState.jobId}?format=xlsx`; xlsx.classList.remove("hidden"); }
    if (csv)  { csv.href  = `/api/fulltext/result/${ftState.jobId}?format=csv`;  csv.classList.remove("hidden"); }
    document.getElementById("ftProgressLabel").textContent = "Done — download below.";
  } else if (status === "cancelled") {
    document.getElementById("ftProgressLabel").textContent = "Cancelled.";
  } else {
    document.getElementById("ftProgressLabel").textContent = "Stopped with error.";
  }
}

async function ftCancelJob() {
  if (!ftState.jobId) return;
  try {
    await fetch(`/api/fulltext/cancel/${ftState.jobId}`, { method: "POST" });
  } catch {}
}

function ftRestart() {
  ftState.jobId = null;
  document.getElementById("ftProgressCard")?.classList.add("hidden");
  document.getElementById("ftLiveLog")?.classList.add("hidden");
  const tog = document.getElementById("ftBtnToggleLog");
  if (tog) tog.textContent = "Show live log";
  ftHideProgressUI();
}

// ── Init ─────────────────────────────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", () => {
  if (!document.getElementById("pane-fulltext")) return;

  // Restore saved field values for the full-text tab
  ["ftStudySynopsis",
   "ftModelSelect", "ftModelSelectClaude", "ftModelSelectGoogle",
   "ftApiKey", "ftApiKeyClaude", "ftApiKeyGoogle",
   "ftVerbosity", "ftReasoningEffort",
   "ftTemperature", "ftTemperatureClaude", "ftTemperatureGoogle"].forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const k = `ft_${id}`;
    const v = localStorage.getItem(k);
    if (v != null) el.value = v;
    const evt = el.tagName === "SELECT" ? "change" : "input";
    el.addEventListener(evt, () => localStorage.setItem(k, el.value));
  });

  const ms = document.getElementById("ftModelSelect");
  if (ms) ms.addEventListener("change", ftToggleParamGroups);

  ftInitProviderSelector();
  ftInitCriteriaLists();
  ftInitUpload();
  ftInitPromptPreview();

  document.getElementById("ftBtnSend")?.addEventListener("click", ftSendJob);
  document.getElementById("ftBtnCancel")?.addEventListener("click", ftCancelJob);
  document.getElementById("ftBtnRestart")?.addEventListener("click", ftRestart);
  document.getElementById("ftBtnToggleLog")?.addEventListener("click", () => {
    const log = document.getElementById("ftLiveLog");
    const btn = document.getElementById("ftBtnToggleLog");
    if (!log || !btn) return;
    const hidden = log.classList.toggle("hidden");
    btn.textContent = hidden ? "Show live log" : "Hide live log";
  });
});
