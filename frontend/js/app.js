// ── State ─────────────────────────────────────────────────────────────────────

const state = {
  workbook: null,
  filename: null,
  selectedSheet: null,
  rows: [],
  provider: "openai",
};

// ── Preferences ───────────────────────────────────────────────────────────────

const savePref = debounce((id, value) => localStorage.setItem(id, value), 200);

[
  "modelSelect", "modelSelectClaude", "modelSelectGoogle",
  "studySynopsis",
  "inclusionCriteria",
  "exclusionCriteria",
  "apiKey", "apiKeyClaude", "apiKeyGoogle",
  "verbosity", "reasoningEffort",
  "temperature", "temperatureClaude", "temperatureGoogle",
  "tierSelect", "tierSelectClaude", "tierSelectGoogle",
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

// ── Dark mode ────────────────────────────────────────────────────────────────

function applyTheme(dark) {
  document.documentElement.setAttribute('data-theme', dark ? 'dark' : 'light');
  const moon = document.getElementById('iconMoon');
  const sun  = document.getElementById('iconSun');
  if (moon) moon.classList.toggle('hidden', dark);
  if (sun)  sun.classList.toggle('hidden', !dark);
}

// ── Init on load ──────────────────────────────────────────────────────────────

document.addEventListener("DOMContentLoaded", () => {
  // Dark mode
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  applyTheme(isDark);
  const btnDark = document.getElementById('btnDarkMode');
  if (btnDark) {
    btnDark.addEventListener('click', () => {
      const nowDark = document.documentElement.getAttribute('data-theme') !== 'dark';
      localStorage.setItem('theme', nowDark ? 'dark' : 'light');
      applyTheme(nowDark);
    });
  }

  // Restore saved field values
  const restoreIds = [
    "modelSelect", "modelSelectClaude", "modelSelectGoogle",
    "studySynopsis",
    "apiKey", "apiKeyClaude", "apiKeyGoogle",
    "verbosity", "reasoningEffort",
    "temperature", "temperatureClaude", "temperatureGoogle",
  ];
  restoreIds.forEach(id => {
    const v = localStorage.getItem(id);
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
  initProviderSelector();
  initAdvancedSettings();
});
