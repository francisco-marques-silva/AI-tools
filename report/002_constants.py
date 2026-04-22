"""
constants.py — Paths, colors and shared configuration.
"""

from pathlib import Path

# ── Paths ────────────────────────────────────────────────────────────
SCRIPT_DIR  = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent
INPUT_DIR   = PROJECT_DIR / "input"
OUTPUT_DIR  = PROJECT_DIR / "output"

# ── Model colours (shared by chart_data + graphic scripts) ───────────
MODEL_COLORS = {
    "gpt-4o":     "#4472C4",
    "gpt-5_2":    "#ED7D31",
    "gpt_5_2":    "#ED7D31",
    "gpt-5-mini": "#A5A5A5",
    "gpt_5_mini": "#A5A5A5",
    "gpt-5-nano": "#FFC000",
    "gpt_5_nano": "#FFC000",
    "gpt_4o":     "#4472C4",
}

DEFAULT_COLORS = [
    "#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#70AD47",
    "#9B59B6", "#E74C3C", "#1ABC9C", "#34495E", "#F39C12",
]

# ── Chart export settings ────────────────────────────────────────────
DPI        = 300
FIG_FORMAT = "png"
FACECOLOR  = "white"
