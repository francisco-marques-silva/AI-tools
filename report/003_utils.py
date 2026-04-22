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
    if d == "included":
        return "include"
    if d == "excluded":
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
