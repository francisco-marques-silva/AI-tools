// ── State ─────────────────────────────────────────────────────────────────────

const state = {
  workbook: null,
  filename: null,
  selectedSheet: null,
  rows: [],
};

// ── Preferences ───────────────────────────────────────────────────────────────

const savePref = debounce((id, value) => localStorage.setItem(id, value), 200);

[
  "modelSelect",
  "studySynopsis",
  "inclusionCriteria",
  "exclusionCriteria",
  "apiKey",
  "verbosity", "reasoningEffort",
  "temperature",
  "tierSelect",
  "adv_concurrent", "adv_concurrent_max", "adv_concurrent_min",
  "adv_record_retries", "adv_aiup_after", "adv_max_retries", "adv_base_backoff",
].forEach(id => {
  const el = document.getElementById(id);
  if (!el) return;
  const evt = el.tagName === "SELECT" ? "change" : "input";
  el.addEventListener(evt, () => savePref(id, el.value));
});

// ── Tab scroll to top ─────────────────────────────────────────────────────────

document.querySelectorAll('.tab-radio').forEach(r => {
  r.addEventListener('change', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
});

// ── Init on load ──────────────────────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", () => {
  const savedModel = localStorage.getItem("modelSelect");
  const savedStudy = localStorage.getItem("studySynopsis");
  const savedKey   = localStorage.getItem("apiKey");

  if (savedModel) { const el = document.getElementById("modelSelect"); if (el) el.value = savedModel; }
  if (savedStudy) { const el = document.getElementById("studySynopsis"); if (el) el.value = savedStudy; }
  if (savedKey)   { const el = document.getElementById("apiKey");        if (el) el.value = savedKey; }

  const paramIds = ["verbosity", "reasoningEffort", "temperature"];
  paramIds.forEach(id => {
    const v  = localStorage.getItem(id);
    const el = document.getElementById(id);
    if (el && v != null) el.value = v;
  });

  const ms = document.getElementById("modelSelect");
  if (ms) {
    ms.addEventListener("change", toggleParamGroups);
    toggleParamGroups();
  }

  addParamTooltips();
  initCriteriaLists();
  initAdvancedSettings();
});
