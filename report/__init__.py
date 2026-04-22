"""
report — Modular AI screening analysis package.

Physical files use numbered prefixes for visual organisation in file
browsers (e.g. 002_constants.py, 003_utils.py).  This __init__ registers
them under their clean import names so that ``from report.utils import …``
and ``from .utils import …`` work transparently.
"""

import importlib as _importlib
import sys as _sys

# (clean_name, numbered_file) — ORDER MATTERS: dependencies first.
_MODULE_MAP = [
    ("constants",        "002_constants"),
    ("utils",            "003_utils"),
    ("file_detection",   "004_file_detection"),
    ("analysis",         "005_analysis"),
    ("chart_data",       "006_chart_data"),
    ("docx_helpers",     "007_docx_helpers"),
    ("report_generator", "008_report_generator"),
    ("fp_workarea",      "009_fp_workarea"),
]

for _clean, _numbered in _MODULE_MAP:
    _mod = _importlib.import_module(f".{_numbered}", __name__)
    _sys.modules[f"{__name__}.{_clean}"] = _mod
    globals()[_clean] = _mod
