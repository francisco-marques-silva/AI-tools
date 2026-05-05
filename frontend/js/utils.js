// ── Shared utilities ──────────────────────────────────────────────────────────

const debounce = (fn, ms = 300) => {
  let t;
  return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
};

function splitLines(txt) {
  return txt.split(/\r?\n/).map(s => s.replace(/^[-•\s]+/, "").trim()).filter(Boolean);
}

// ── CSV helpers ───────────────────────────────────────────────────────────────

function csvToArray(text) {
  const lines = text.split(/\r?\n/).filter(l => l.length > 0);
  if (!lines.length) return [];
  const delim = detectDelimiter(lines[0]);
  if (lines[0].charCodeAt(0) === 0xfeff) lines[0] = lines[0].slice(1);
  const header = splitCsvLine(lines[0], delim).map(s => s.trim());
  const out = [];
  for (let i = 1; i < lines.length; i++) {
    const cols = splitCsvLine(lines[i], delim);
    const obj = {};
    header.forEach((h, idx) => obj[h] = cols[idx] ?? "");
    out.push(obj);
  }
  return out;
}

function detectDelimiter(headerLine) {
  const commas = (headerLine.match(/,/g) || []).length;
  const semis  = (headerLine.match(/;/g) || []).length;
  return semis > commas ? ";" : ",";
}

function splitCsvLine(line, delim) {
  const out = []; let cur = ""; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else { inQ = !inQ; }
      continue;
    }
    if (c === delim && !inQ) { out.push(cur); cur = ""; continue; }
    cur += c;
  }
  out.push(cur);
  return out.map(s => s.trim());
}

function decodeTextWithFallback(buffer) {
  const dec = (enc) => new TextDecoder(enc, { fatal: false }).decode(buffer);
  let txt = dec("utf-8");
  const bad = (txt.match(/�/g) || []).length;
  if (bad > 0 && bad / Math.max(1, txt.length) > 0.001) {
    try { txt = dec("windows-1252"); } catch {}
  }
  return txt;
}

// ── Column inference ──────────────────────────────────────────────────────────
// Returns { cols: string[], rows: object[] }
// cols = column names after normalisation (always includes "title"/"abstract" if found)
// rows = normalised rows (original rows if no mapping needed)

function inferColumns(rows) {
  if (!rows.length) return { cols: [], rows: [] };
  const rawCols = Object.keys(rows[0]);
  const map = rawCols.reduce((a, c) => { a[c.toLowerCase().trim()] = c; return a; }, {});
  const variants = {
    title:    ["title", "título", "titulo"],
    abstract: ["abstract", "resumo", "summary"],
  };
  const chosen = {};
  for (const k in variants) {
    for (const v of variants[k]) {
      if (map[v]) { chosen[k] = map[v]; break; }
    }
  }
  if (chosen.title || chosen.abstract) {
    const normalized = rows.map(row => {
      const r2 = { ...row };
      if (chosen.title)    r2.title    = row[chosen.title];
      if (chosen.abstract) r2.abstract = row[chosen.abstract];
      return r2;
    });
    return { cols: Object.keys(normalized[0] ?? {}), rows: normalized };
  }
  return { cols: rawCols, rows };
}
