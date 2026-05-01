"""
fp_workarea.py — False Positives Workspace.

Generates fp_workarea_YYYYMMDD_HHMMSS.xlsx in the output/ folder alongside the
Word report and chart data.  One sheet per study (e.g. MINO, NMDA, ZEBRA).

Each sheet lists every article flagged as a false positive in at least one AI
run, sorted by consensus count (most-agreed FPs first).

Columns:
    Título          Article title
    Resumo          Article abstract
    Quantas_IA_FP   Number of AI runs that classified the article as FP
    IAs             Which runs, e.g. "gpt-4o (1º teste), gpt-5_2 (2º teste)"

Definition of false positive (mirrors run_diagnostic in analysis.py):
    FP = AI decision "maybe"  AND  Human decision "exclude"

Integration: called automatically by main.py via generate_fp_workarea().
Standalone:  python backend/report/fp_workarea.py
"""

from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment

from .utils import normalise_title


# ── Column widths ─────────────────────────────────────────────────────────────

_COL_WIDTHS = {"Título": 60, "Resumo": 80, "Quantas_IA_FP": 18, "IAs": 50}


# ── Excel formatting ──────────────────────────────────────────────────────────

def _format_sheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame):
    ws = writer.sheets[sheet_name]
    for i, col in enumerate(df.columns, start=1):
        ws.column_dimensions[chr(64 + i)].width = _COL_WIDTHS.get(col, 20)
    wrap_top = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = wrap_top


# ── Core builder ──────────────────────────────────────────────────────────────

def _build_sheet_df(pn: str, projects: dict, fp_results: dict) -> pd.DataFrame:
    """
    Aggregate false positive articles for one project across all AI runs.

    Returns a DataFrame ready to write as an Excel sheet, sorted by
    Quantas_IA_FP descending.
    """
    fp_data = {}   # normalised_title_key -> {title, abstract, ia_labels}

    for mn in sorted(fp_results[pn]):
        model_name = projects[pn]["models"][mn]["name"]
        for tn in sorted(fp_results[pn][mn]):
            label = f"{model_name} ({tn}º teste)"
            articles = fp_results[pn][mn][tn].get("fp_articles", [])
            for art in articles:
                key = normalise_title(art.get("title", ""))
                if not key:
                    continue
                if key not in fp_data:
                    fp_data[key] = {
                        "title":     art.get("title", ""),
                        "abstract":  art.get("abstract", "") or "",
                        "ia_labels": [],
                    }
                fp_data[key]["ia_labels"].append(label)

    if not fp_data:
        return pd.DataFrame(columns=["Título", "Resumo", "Quantas_IA_FP", "IAs"])

    rows = [
        {
            "Título":       d["title"],
            "Resumo":       d["abstract"],
            "Quantas_IA_FP": len(d["ia_labels"]),
            "IAs":          ", ".join(d["ia_labels"]),
        }
        for d in fp_data.values()
    ]

    return (
        pd.DataFrame(rows)
        .sort_values("Quantas_IA_FP", ascending=False)
        .reset_index(drop=True)
    )


# ── Public API (called by main.py) ────────────────────────────────────────────

def generate_fp_workarea(projects: dict, all_results: dict, output_dir: Path) -> Path:
    """
    Build and write the FP workspace Excel file.

    Parameters
    ----------
    projects    Project structure from build_project_structure()
    all_results Results dict from run_all_analyses(); must contain
                'false_positives' key populated by run_diagnostic()
    output_dir  Destination folder (same as Word report + chart XLSX)

    Returns
    -------
    Path to the generated .xlsx file
    """
    fp_results = all_results.get("false_positives", {})
    sheets = {}

    for pn in sorted(projects):
        if pn not in fp_results:
            continue
        proj_name  = projects[pn]["name"].upper()
        sheet_df   = _build_sheet_df(pn, projects, fp_results)
        n_fp       = len(sheet_df)
        print(f"    {proj_name}: {n_fp} unique FP articles")
        sheets[proj_name] = sheet_df

    if not sheets:
        print("    No FP data to write.")
        return None

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path  = output_dir / f"fp_workarea_{timestamp}.xlsx"

    with pd.ExcelWriter(str(out_path), engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            _format_sheet(writer, sheet_name, df)

    return out_path


# ── Standalone entry-point ────────────────────────────────────────────────────

def _run_standalone():
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
    from backend.report.constants import INPUT_DIR, OUTPUT_DIR
    from backend.report.file_detection import scan_input_dir, build_project_structure
    from backend.report.analysis import run_diagnostic

    print(f"\nScanning: {INPUT_DIR}")
    ai_files, human_files, meta_path = scan_input_dir(INPUT_DIR)
    projects, _ = build_project_structure(ai_files, human_files, meta_path)

    fp_results = {}

    for pn in sorted(projects):
        proj = projects[pn]
        if not proj["human_tiab"]:
            print(f"  {proj['name']}: no human TIAB, skipping.")
            continue

        fp_results[pn] = {}
        for mn in sorted(proj["models"]):
            fp_results[pn][mn] = {}
            model = proj["models"][mn]
            for tn, test_info in sorted(model["tests"].items()):
                try:
                    r = run_diagnostic(test_info["path"], proj["human_tiab"])
                    if r:
                        fp_results[pn][mn][tn] = {
                            "fp_articles": r.get("fp_articles", []),
                        }
                except Exception as exc:
                    print(f"    WARNING {proj['name']} / {model['name']} / test {tn}: {exc}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    all_results = {"false_positives": fp_results}
    path = generate_fp_workarea(projects, all_results, OUTPUT_DIR)
    if path:
        print(f"\nOutput saved to: {path}")


if __name__ == "__main__":
    _run_standalone()
