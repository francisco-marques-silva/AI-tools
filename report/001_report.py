"""
001_report.py — Thin CLI entry-point.

Orchestrates the full pipeline:
    scan input/  →  build project structure  →  validate data
    →  run all analyses  →  generate Word report  →  export chart XLSX

Usage:
    python report/001_report.py                       # auto-detect input/
    python report/001_report.py --input_dir DIR       # custom input folder

Modules consumed (via report package aliases):
    constants        Paths and shared configuration
    utils            I/O, normalisation, formatting helpers
    file_detection   File pattern matching and project grouping
    analysis         Statistical analysis functions
    chart_data       Chart XLSX export  (data_grafics_*.xlsx)
    docx_helpers     Word formatting primitives
    report_generator Report document builder (tables only)
"""

import argparse
import sys
from pathlib import Path

# Ensure the project root is on sys.path so ``from report.*`` works
# even when this script is invoked directly: python report/001_report.py
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from report.constants import INPUT_DIR, OUTPUT_DIR
from report.utils import fmt, fmt_pct, normalise_model_name
from report.file_detection import scan_input_dir, build_project_structure
from report.analysis import (
    run_diagnostic,
    run_fulltext_check,
    run_listfinal_check,
    run_test_retest,
)
from report.chart_data import export_chart_data
from report.report_generator import generate_report
from report.fp_workarea import generate_fp_workarea


# ===========================================================================
#  DATA VALIDATION
# ===========================================================================

def validate_data(projects, metadados):
    """Check correspondence between files on disk and metadata spreadsheet."""
    issues = []

    for pn, proj in projects.items():
        if not proj["human_tiab"]:
            issues.append(
                f"Project '{proj['name']}': no human TIAB file. "
                "Diagnostic analysis will not be possible (test-retest only)."
            )
        if not proj["human_fulltext"]:
            issues.append(
                f"Project '{proj['name']}': no human Fulltext file. "
                "Fulltext verification will not be possible."
            )
        if not proj.get("human_listfinal"):
            issues.append(
                f"Project '{proj['name']}': no Listfinal file. "
                "Listfinal verification will not be possible."
            )

        for mn, model in proj["models"].items():
            tests = model["tests"]
            if 1 not in tests or 2 not in tests:
                missing = [t for t in [1, 2] if t not in tests]
                issues.append(
                    f"Project '{proj['name']}', model '{model['name']}': "
                    f"missing test {'st, '.join(str(m) for m in missing)}."
                )

    if metadados is not None:
        meta_codes = set(metadados["code"].astype(str).tolist())
        ai_codes = set()
        for proj in projects.values():
            for model in proj["models"].values():
                for test in model["tests"].values():
                    ai_codes.add(test["code"])

        missing_in_meta = ai_codes - meta_codes
        missing_in_files = meta_codes - ai_codes

        if missing_in_meta:
            issues.append(
                f"Codes present in AI files but missing from metadata: "
                f"{', '.join(sorted(missing_in_meta))}"
            )
        if missing_in_files:
            issues.append(
                f"Codes present in metadata but without corresponding AI file: "
                f"{', '.join(sorted(missing_in_files))}"
            )

    if metadados is not None:
        for pn, proj in projects.items():
            for mn, model in proj["models"].items():
                for tn2, test in model["tests"].items():
                    code = test["code"]
                    meta_match = metadados[metadados["code"].astype(str) == str(code)]
                    if not meta_match.empty:
                        meta_model = meta_match.iloc[0]["model"]
                        file_model = model["name"]
                        if normalise_model_name(str(meta_model)) != normalise_model_name(file_model):
                            issues.append(
                                f"Model name inconsistency for code {code}: "
                                f"metadata='{meta_model}', file='{file_model}'"
                            )

    return issues


# ===========================================================================
#  RUN ALL ANALYSES
# ===========================================================================

def run_all_analyses(projects, metadados):
    """Execute every analysis and return a results dict."""
    all_results = {}

    # ---- Validation ----
    issues = validate_data(projects, metadados)
    all_results["validation_issues"] = issues

    print("\n" + "=" * 70)
    print("  UNIFIED REPORT — Processing")
    print("=" * 70)

    if issues:
        print("\n  ⚠ Validation notes:")
        for issue in issues:
            print(f"    • {issue}")
    else:
        print("\n  ✓ All data validated successfully.")

    # ---- Diagnostic analysis ----
    print("\n  Running diagnostic analyses...")
    diag_results = {}
    fn_results = {}
    fp_results = {}

    for pn in sorted(projects.keys()):
        proj = projects[pn]
        if not proj["human_tiab"]:
            print(f"    {proj['name']}: no human TIAB, skipping diagnostic.")
            continue

        diag_results[pn] = {}
        fn_results[pn] = {}
        fp_results[pn] = {}

        for mn in sorted(proj["models"].keys()):
            model = proj["models"][mn]
            diag_results[pn][mn] = {}
            fn_results[pn][mn] = {}
            fp_results[pn][mn] = {}

            for test_num, test_info in sorted(model["tests"].items()):
                print(f"    {proj['name']} / {model['name']} / test {test_num}...", end=" ")
                try:
                    r = run_diagnostic(test_info["path"], proj["human_tiab"])
                    diag_results[pn][mn][test_num] = r
                    if r:
                        fn_results[pn][mn][test_num] = {
                            "fn": r["fn"],
                            "n_paired": r["n_paired"],
                            "fn_titles": r.get("fn_titles", []),
                        }
                        fp_results[pn][mn][test_num] = {
                            "fp": r["fp"],
                            "tn": r["tn"],
                            "n_paired": r["n_paired"],
                            "fp_titles": r.get("fp_titles", []),
                            "fp_articles": r.get("fp_articles", []),
                        }
                        print(
                            f"OK (Sens={fmt_pct(r['metrics']['Sensitivity'])}, "
                            f"Kappa={fmt(r['kappa'], 3)})"
                        )
                    else:
                        print("No valid data.")
                except Exception as e:
                    print(f"ERROR: {e}")
                    diag_results[pn][mn][test_num] = None

    all_results["diagnostic"] = diag_results
    all_results["false_negatives"] = fn_results
    all_results["false_positives"] = fp_results

    # ---- Fulltext Check ----
    print("\n  Running fulltext verification...")
    ft_results = {}

    for pn in sorted(projects.keys()):
        proj = projects[pn]
        if not proj["human_fulltext"]:
            print(f"    {proj['name']}: no human fulltext, skipping.")
            continue

        ft_results[pn] = {}
        for mn in sorted(proj["models"].keys()):
            model = proj["models"][mn]
            ft_results[pn][mn] = {}
            for test_num, test_info in sorted(model["tests"].items()):
                print(f"    {proj['name']} / {model['name']} / test {test_num}...", end=" ")
                try:
                    r = run_fulltext_check(test_info["path"], proj["human_fulltext"])
                    ft_results[pn][mn][test_num] = r
                    print(
                        f"OK (Capture={fmt_pct(r['capture_rate'])}, "
                        f"Missed={r['n_missed']})"
                    )
                except Exception as e:
                    print(f"ERROR: {e}")

    all_results["fulltext"] = ft_results

    # ---- Listfinal Check ----
    print("\n  Running Listfinal verification...")
    lf_results = {}

    for pn in sorted(projects.keys()):
        proj = projects[pn]
        if not proj.get("human_listfinal"):
            print(f"    {proj['name']}: no Listfinal file, skipping.")
            continue

        lf_results[pn] = {}
        for mn in sorted(proj["models"].keys()):
            model = proj["models"][mn]
            lf_results[pn][mn] = {}
            for test_num, test_info in sorted(model["tests"].items()):
                print(f"    {proj['name']} / {model['name']} / test {test_num}...", end=" ")
                try:
                    r = run_listfinal_check(test_info["path"], proj["human_listfinal"])
                    lf_results[pn][mn][test_num] = r
                    print(
                        f"OK (Capture={fmt_pct(r['capture_rate'])}, "
                        f"Missed={r['n_missed']})"
                    )
                except Exception as e:
                    print(f"ERROR: {e}")

    all_results["listfinal"] = lf_results

    # ---- Test-Retest ----
    print("\n  Running test-retest...")
    tr_results = {}

    for pn in sorted(projects.keys()):
        proj = projects[pn]
        tr_results[pn] = {}

        for mn in sorted(proj["models"].keys()):
            model = proj["models"][mn]
            if 1 in model["tests"] and 2 in model["tests"]:
                print(f"    {proj['name']} / {model['name']}...", end=" ")
                try:
                    r = run_test_retest(
                        model["tests"][1]["path"],
                        model["tests"][2]["path"],
                    )
                    tr_results[pn][mn] = r
                    print(
                        f"OK (Agree={fmt_pct(r['binary_pct'])}, "
                        f"Kappa={fmt(r['kappa'], 3)})"
                    )
                except Exception as e:
                    print(f"ERROR: {e}")

    all_results["test_retest"] = tr_results

    return all_results


# ===========================================================================
#  MAIN
# ===========================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Generates a unified report of all AI screening analyses.",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--input_dir", "-i", default=None,
        help="Folder with input files (default: input/).",
    )
    parser.add_argument(
        "--output_dir", "-o", default=None,
        help="Folder for output files (default: output/).",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir) if args.input_dir else INPUT_DIR
    output_dir = Path(args.output_dir) if args.output_dir else OUTPUT_DIR
    if not input_dir.is_dir():
        print(f"\n  ERROR: Input folder not found: {input_dir}")
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    # ---- Detect files ----
    print(f"\n  Scanning folder: {input_dir}")
    ai_files, human_files, metadados_path = scan_input_dir(input_dir)

    print(f"  AI files found:      {len(ai_files)}")
    print(f"  Human files found:   {len(human_files)}")
    print(f"  Metadata found:      {'Yes' if metadados_path else 'No'}")

    if not ai_files:
        print("\n  ERROR: No AI files found in input/.")
        print("  Expected: YYYYMMDD - model - Xº teste - project.xlsx")
        sys.exit(1)

    # ---- Build structure ----
    projects, metadados = build_project_structure(ai_files, human_files, metadados_path)

    print(f"\n  Projects identified: {len(projects)}")
    for pn in sorted(projects.keys()):
        proj = projects[pn]
        n_models = len(proj["models"])
        models_list = ", ".join(m["name"] for m in proj["models"].values())
        print(f"    • {proj['name']}: {n_models} models ({models_list})")

    # ---- Run analyses ----
    all_results = run_all_analyses(projects, metadados)

    # ---- Generate Word report (tables only, no charts) ----
    print("\n  Generating Word report...")
    docx_path = generate_report(projects, metadados, all_results, output_dir)
    print(f"\n  ✓ Report generated: {docx_path.name}")

    # ---- Export chart data XLSX ----
    print("  Exporting chart data XLSX...")
    xlsx_path = export_chart_data(projects, metadados, all_results, output_dir)
    print(f"  ✓ Chart data exported: {xlsx_path.name}")

    # ---- Generate FP workspace XLSX ----
    print("  Generating FP workspace XLSX...")
    fp_path = generate_fp_workarea(projects, all_results, output_dir)
    if fp_path:
        print(f"  ✓ FP workspace exported: {fp_path.name}")

    print(f"\n    Output folder: {output_dir}")
    print("=" * 70)
    print()


if __name__ == "__main__":
    main()
