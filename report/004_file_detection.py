"""
file_detection.py — Scan the input/ folder and build the project structure.
"""

import re
from pathlib import Path

import pandas as pd

from .utils import normalise_model_name

# ── Filename patterns ────────────────────────────────────────────────

# AI file:  YYYYMMDD - model - Xo teste - project.xlsx
AI_FILE_PATTERN = re.compile(
    r"^(\d{8})\s*-\s*(.+?)\s*-\s*(\d)[ºo°]\s*teste\s*-\s*(.+)\.xlsx$",
    re.IGNORECASE,
)

# Human file:  Project - TIAB.xlsx | Project - Fulltext.xlsx | Project - Listfinal.xlsx
HUMAN_FILE_PATTERN = re.compile(
    r"^(.+?)\s*-\s*(TIAB|Fulltext|Listfinal)\.xlsx$",
    re.IGNORECASE,
)


# ── Parsing helpers ──────────────────────────────────────────────────

def parse_ai_filename(filename: str):
    """Return dict {code, model, test_num, project} or None."""
    m = AI_FILE_PATTERN.match(filename)
    if not m:
        return None
    return {
        "code": m.group(1),
        "model": m.group(2).strip(),
        "model_norm": normalise_model_name(m.group(2).strip()),
        "test_num": int(m.group(3)),
        "project": m.group(4).strip(),
        "project_norm": m.group(4).strip().lower(),
    }


def parse_human_filename(filename: str):
    """Return dict {project, type} or None.  type = 'tiab' | 'fulltext' | 'listfinal'"""
    m = HUMAN_FILE_PATTERN.match(filename)
    if not m:
        return None
    return {
        "project": m.group(1).strip(),
        "project_norm": m.group(1).strip().lower(),
        "type": m.group(2).strip().lower(),
    }


# ── Scan & structure ─────────────────────────────────────────────────

def scan_input_dir(input_dir: Path):
    """Scan input/ and return (ai_files, human_files, metadados_path)."""
    ai_files = []
    human_files = []
    metadados_path = None

    for f in sorted(input_dir.iterdir()):
        if not f.is_file():
            continue
        name = f.name

        if name.lower() == "metadata.xlsx":
            metadados_path = f
            continue

        parsed_ai = parse_ai_filename(name)
        if parsed_ai:
            parsed_ai["path"] = f
            ai_files.append(parsed_ai)
            continue

        parsed_human = parse_human_filename(name)
        if parsed_human:
            parsed_human["path"] = f
            human_files.append(parsed_human)
            continue

    return ai_files, human_files, metadados_path


def build_project_structure(ai_files, human_files, metadados_path):
    """
    Organise data into per-project structure:

    {
      project_norm: {
        'name': str,
        'human_tiab': Path | None,
        'human_fulltext': Path | None,
        'human_listfinal': Path | None,
        'models': {
          model_norm: {
            'name': str,
            'tests': {1: {path, code}, 2: {path, code}}
          }
        }
      }
    }
    """
    projects = {}

    for ai in ai_files:
        pn = ai["project_norm"]
        mn = ai["model_norm"]
        if pn not in projects:
            projects[pn] = {
                "name": ai["project"],
                "human_tiab": None,
                "human_fulltext": None,
                "human_listfinal": None,
                "models": {},
            }
        if mn not in projects[pn]["models"]:
            projects[pn]["models"][mn] = {
                "name": ai["model"],
                "tests": {},
            }
        projects[pn]["models"][mn]["tests"][ai["test_num"]] = {
            "path": ai["path"],
            "code": ai["code"],
        }

    for hf in human_files:
        pn = hf["project_norm"]
        if pn not in projects:
            projects[pn] = {
                "name": hf["project"],
                "human_tiab": None,
                "human_fulltext": None,
                "human_listfinal": None,
                "models": {},
            }
        if hf["type"] == "tiab":
            projects[pn]["human_tiab"] = hf["path"]
        elif hf["type"] == "fulltext":
            projects[pn]["human_fulltext"] = hf["path"]
        elif hf["type"] == "listfinal":
            projects[pn]["human_listfinal"] = hf["path"]

    # Metadata
    metadados = None
    if metadados_path and metadados_path.is_file():
        metadados = pd.read_excel(metadados_path)
        if "project" in metadados.columns:
            metadados = metadados.dropna(subset=["project"]).reset_index(drop=True)

    return projects, metadados
