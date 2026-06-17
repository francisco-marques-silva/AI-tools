"""
utils.py — General I/O, normalisation and formatting helpers.
"""

import re
from pathlib import Path

import numpy as np
import pandas as pd


# ── File I/O ─────────────────────────────────────────────────────────

def load_file(path: str) -> pd.DataFrame:
    """Read CSV or Excel into a DataFrame."""
    ext = Path(path).suffix.lower()
    if ext == ".csv":
        try:
            return pd.read_csv(path, encoding="utf-8")
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="latin-1")
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    else:
        raise ValueError(f"Unsupported format: {ext}")


def normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Lower-case and strip column names."""
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    return df


# ── Title / decision normalisation ───────────────────────────────────

_HTML_TAG_RE = re.compile(r"<[^>]+")
_NONALNUM_RE = re.compile(r"[^a-z0-9]")


def normalise_title(s) -> str:
    """Key used for pairing articles by title.

    Strips HTML tags, then removes all non-alphanumeric characters (lowercase).
    Handles D<sub>1</sub> vs D(1) vs D1, brackets, extra whitespace, etc.
    """
    if pd.isna(s):
        return ""
    s = _HTML_TAG_RE.sub("", str(s))
    return _NONALNUM_RE.sub("", s.lower())


def normalise_decision(s) -> str:
    if pd.isna(s):
        return ""
    d = str(s).strip().lower()
    _INCLUDE = {
        "include", "included", "incluir", "incluído", "incluido",
        "incluso", "sim", "yes", "1", "true", "maybe", "talvez",
    }
    _EXCLUDE = {
        "exclude", "excluded", "excluir", "excluído", "excluido",
        "não", "nao", "no", "0", "false",
    }
    if d in _INCLUDE:
        return "include"
    if d in _EXCLUDE:
        return "exclude"
    return d


def binarise_decision(d: str) -> str:
    """include/maybe -> maybe  |  exclude -> exclude"""
    d = normalise_decision(d)
    if d in ("include", "maybe"):
        return "maybe"
    return d


def normalise_model_name(name: str) -> str:
    """Standardise model name variants: gpt-5-2, gpt-5.2, gpt-5_2 -> gpt_5_2"""
    return name.strip().lower().replace(".", "_").replace("-", "_").replace(" ", "_")


# ── Numeric formatting ───────────────────────────────────────────────

def fmt(v, d=4):
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        if np.isnan(v):
            return "N/A"
        if np.isinf(v):
            return "∞"
        return f"{v:.{d}f}"
    return str(v)


def fmt_pct(v):
    if isinstance(v, float) and not np.isnan(v) and not np.isinf(v):
        return f"{v * 100:.1f}%"
    return "-"


# ── F1 vs Listfinal (gold standard) ──────────────────────────────────

def compute_f1_lf(d, l):
    """F1 score against the Listfinal gold standard (not TIAB).

    tp_lf = LF articles captured by AI; fn_lf = LF articles missed by AI;
    fp_lf = AI positives that are not in LF (overinclusion).
    Returns NaN if either input is missing.
    """
    if d is None or l is None:
        return float("nan")
    ai_pos = d["tp"] + d["fp"]
    tp_lf = l["n_captured"]
    fp_lf = max(0, ai_pos - tp_lf)
    fn_lf = l["n_missed"]
    denom = 2 * tp_lf + fp_lf + fn_lf
    return (2 * tp_lf / denom) if denom > 0 else float("nan")


def compute_metrics_vs_lf(n_universe, n_positives, lf):
    """Sens/Spec/F1 of any decision-set against the Listfinal gold standard.

    Works for both the AI and the human reviewer.

    n_universe : int
        Total articles in the screening universe (paired/found).
    n_positives : int
        Number of articles the decision-set marked as include/maybe.
    lf : dict
        Output of `run_listfinal_check` (or `run_human_vs_lf`); must contain
        `n_captured` (TP_lf), `n_missed` (FN_lf), `n_found` (LF-in-universe).

    Returns dict with tp/fp/fn/tn vs LF, sens/spec/f1, and inclusion_rate.
    Returns NaN-filled metrics when data are missing or inconsistent.
    """
    nan_result = {
        "tp_lf": 0, "fp_lf": 0, "fn_lf": 0, "tn_lf": 0,
        "sens_lf": float("nan"), "spec_lf": float("nan"),
        "f1_lf": float("nan"), "inclusion_rate": float("nan"),
        "n_universe": n_universe or 0, "n_positives": n_positives or 0,
    }
    if lf is None or n_universe is None or n_universe <= 0:
        return nan_result
    tp = int(lf.get("n_captured", 0))
    fn = int(lf.get("n_missed", 0))
    lf_in = tp + fn
    fp = max(0, int(n_positives or 0) - tp)
    tn = max(0, int(n_universe) - lf_in - fp)
    sens = tp / (tp + fn) if (tp + fn) else float("nan")
    spec = tn / (tn + fp) if (tn + fp) else float("nan")
    f1 = 2 * tp / (2 * tp + fp + fn) if (2 * tp + fp + fn) else float("nan")
    inclusion_rate = (n_positives / n_universe) if n_universe else float("nan")
    return {
        "tp_lf": tp, "fp_lf": fp, "fn_lf": fn, "tn_lf": tn,
        "sens_lf": sens, "spec_lf": spec, "f1_lf": f1,
        "inclusion_rate": inclusion_rate,
        "n_universe": n_universe, "n_positives": n_positives or 0,
    }
