"""
analysis.py — Statistical analysis functions (diagnostic, fulltext, listfinal, test-retest).
"""

from pathlib import Path

import numpy as np
import pandas as pd

from .utils import (
    load_file,
    normalise_columns,
    normalise_title,
    normalise_decision,
    binarise_decision,
)


# ── Core statistical helpers ─────────────────────────────────────────

def confusion_matrix(ai_decisions, human_decisions):
    """Compute TP, FP, FN, TN (gold standard = human)."""
    tp = int(((ai_decisions == "maybe")   & (human_decisions == "maybe")).sum())
    fp = int(((ai_decisions == "maybe")   & (human_decisions == "exclude")).sum())
    fn = int(((ai_decisions == "exclude") & (human_decisions == "maybe")).sum())
    tn = int(((ai_decisions == "exclude") & (human_decisions == "exclude")).sum())
    return tp, fp, fn, tn


def calc_metrics(tp, fp, fn, tn):
    n    = tp + fp + fn + tn
    sens = tp / (tp + fn) if (tp + fn) else float("nan")
    spec = tn / (tn + fp) if (tn + fp) else float("nan")
    ppv  = tp / (tp + fp) if (tp + fp) else float("nan")
    npv  = tn / (tn + fn) if (tn + fn) else float("nan")
    acc  = (tp + tn) / n  if n else float("nan")
    prev = (tp + fn) / n  if n else float("nan")
    f1   = 2 * tp / (2 * tp + fp + fn) if (2 * tp + fp + fn) else float("nan")
    youd = sens + spec - 1 if not (np.isnan(sens) or np.isnan(spec)) else float("nan")
    lr_p = sens / (1 - spec) if (1 - spec) > 0 else float("inf")
    lr_n = (1 - sens) / spec if spec > 0 else float("inf")
    return {
        "N": n,
        "Prevalence": prev,
        "Sensitivity": sens,
        "Specificity": spec,
        "PPV (Precision)": ppv,
        "NPV": npv,
        "Accuracy": acc,
        "F1 Score": f1,
        "LR+": lr_p,
        "LR-": lr_n,
        "Youden Index": youd,
    }


def calc_kappa(tp, fp, fn, tn):
    n = tp + fp + fn + tn
    if n == 0:
        return float("nan"), float("nan"), float("nan"), float("nan"), ""
    po = (tp + tn) / n
    pe = ((tp + fp) * (tp + fn) + (fn + tn) * (fp + tn)) / (n * n)
    if pe == 1:
        k = 1.0
    else:
        k = (po - pe) / (1 - pe)
    if (1 - pe) == 0:
        se = float("nan")
    else:
        se = np.sqrt(pe * (1 - pe) / (n * (1 - pe) ** 2))
    ci_lo = k - 1.96 * se
    ci_hi = k + 1.96 * se
    if k < 0:       interp = "Poor (< 0)"
    elif k < 0.20:  interp = "Slight (0.00–0.20)"
    elif k < 0.40:  interp = "Fair (0.21–0.40)"
    elif k < 0.60:  interp = "Moderate (0.41–0.60)"
    elif k < 0.80:  interp = "Substantial (0.61–0.80)"
    else:            interp = "Almost Perfect (0.81–1.00)"
    return k, se, ci_lo, ci_hi, interp


# ── Pairing ──────────────────────────────────────────────────────────

def do_pairing(ai_path: Path, human_path: Path):
    """
    Pair AI with human by normalised title.
    Returns (df_paired, n_unpaired_ai, n_unpaired_human).
    """
    ai_df = normalise_columns(load_file(str(ai_path)))
    hu_df = normalise_columns(load_file(str(human_path)))

    ai_df["_title_key"] = ai_df["title"].apply(normalise_title)
    hu_df["_title_key"] = hu_df["title"].apply(normalise_title)

    ai_df["_occ"] = ai_df.groupby("_title_key").cumcount().astype(str)
    hu_df["_occ"] = hu_df.groupby("_title_key").cumcount().astype(str)
    ai_df["_merge_key"] = ai_df["_title_key"] + "__" + ai_df["_occ"]
    hu_df["_merge_key"] = hu_df["_title_key"] + "__" + hu_df["_occ"]

    if "decision" in hu_df.columns:
        hu_df = hu_df.rename(columns={"decision": "decision_human"})

    if "abstract" not in ai_df.columns:
        ai_df["abstract"] = ""

    merged = pd.merge(
        ai_df[["title", "abstract", "_merge_key", "screening_decision"]],
        hu_df[["_merge_key", "decision_human"]],
        on="_merge_key",
        how="outer",
    )

    n_ai_only = merged["screening_decision"].notna() & merged["decision_human"].isna()
    n_hu_only = merged["decision_human"].notna() & merged["screening_decision"].isna()

    paired = merged.dropna(subset=["screening_decision", "decision_human"]).copy()
    paired["ai_bin"] = paired["screening_decision"].apply(binarise_decision)
    paired["hu_bin"] = paired["decision_human"].apply(binarise_decision)

    return paired, int(n_ai_only.sum()), int(n_hu_only.sum())


# ── Diagnostic analysis ──────────────────────────────────────────────

def run_diagnostic(ai_path: Path, human_tiab_path: Path):
    """Full diagnostic analysis.  Returns dict with results or None."""
    paired, n_unpaired_ai, n_unpaired_hu = do_pairing(ai_path, human_tiab_path)

    valid = paired["ai_bin"].isin(["maybe", "exclude"]) & \
            paired["hu_bin"].isin(["maybe", "exclude"])
    paired = paired[valid].copy()

    if paired.empty:
        return None

    tp, fp, fn, tn = confusion_matrix(paired["ai_bin"], paired["hu_bin"])
    metrics = calc_metrics(tp, fp, fn, tn)
    k, se, ci_lo, ci_hi, interp = calc_kappa(tp, fp, fn, tn)

    fp_mask = (paired["ai_bin"] == "maybe") & (paired["hu_bin"] == "exclude")
    fn_mask = (paired["ai_bin"] == "exclude") & (paired["hu_bin"] == "maybe")

    fp_titles = paired[fp_mask]["title"].dropna().tolist()
    fn_titles = paired[fn_mask]["title"].dropna().tolist()
    fp_articles = paired[fp_mask][["title", "abstract"]].to_dict("records")

    return {
        "n_paired": len(paired),
        "n_unpaired_ai": n_unpaired_ai,
        "n_unpaired_hu": n_unpaired_hu,
        "tp": tp, "fp": fp, "fn": fn, "tn": tn,
        "fp_titles": fp_titles,
        "fp_articles": fp_articles,
        "fn_titles": fn_titles,
        "metrics": metrics,
        "kappa": k, "kappa_se": se,
        "kappa_ci_lo": ci_lo, "kappa_ci_hi": ci_hi,
        "kappa_interp": interp,
    }


# ── Fulltext check ───────────────────────────────────────────────────

def run_fulltext_check(ai_path: Path, fulltext_path: Path):
    """Test whether fulltext articles would have been retained by the AI."""
    ai_df = normalise_columns(load_file(str(ai_path)))
    ft_df = normalise_columns(load_file(str(fulltext_path)))

    ai_df["_title_key"] = ai_df["title"].apply(normalise_title)
    ft_df["_title_key"] = ft_df["title"].apply(normalise_title)

    results = []
    for _, row in ft_df.iterrows():
        tk = normalise_title(row["title"])
        abstract = "" if pd.isna(row.get("abstract", "")) else str(row.get("abstract", "")).strip()
        match = ai_df[ai_df["_title_key"] == tk]
        if match.empty:
            results.append({"title": row["title"], "abstract": abstract,
                            "found": False, "ai_decision": "not_found"})
        else:
            ai_dec = normalise_decision(match.iloc[0]["screening_decision"])
            captured = ai_dec in ("maybe", "include")
            results.append({
                "title": row["title"],
                "abstract": abstract,
                "found": True,
                "ai_decision": ai_dec,
                "captured": captured,
            })

    df_results = pd.DataFrame(results)
    n_total = len(df_results)
    n_found = df_results["found"].sum()
    n_not_found = n_total - n_found

    found_df = df_results[df_results["found"]].copy()
    if not found_df.empty:
        found_df["captured"] = found_df["captured"].astype(bool)
    n_captured = int(found_df["captured"].sum()) if not found_df.empty else 0
    n_missed = int((~found_df["captured"]).sum()) if not found_df.empty else 0
    capture_rate = n_captured / n_found if n_found > 0 else float("nan")
    miss_rate = n_missed / n_found if n_found > 0 else float("nan")

    missed_titles = found_df[~found_df["captured"]]["title"].tolist() if not found_df.empty else []
    missed_articles = (
        found_df[~found_df["captured"]][["title", "abstract"]].to_dict("records")
        if not found_df.empty else []
    )

    return {
        "n_fulltext": n_total,
        "n_found": int(n_found),
        "n_not_found": int(n_not_found),
        "n_captured": n_captured,
        "n_missed": n_missed,
        "capture_rate": capture_rate,
        "miss_rate": miss_rate,
        "missed_titles": missed_titles,
        "missed_articles": missed_articles,
    }


# ── Listfinal check (true gold standard) ────────────────────────────

def run_listfinal_check(ai_path: Path, listfinal_path: Path):
    """Test whether final included articles would have been retained by the AI."""
    ai_df = normalise_columns(load_file(str(ai_path)))
    lf_df = normalise_columns(load_file(str(listfinal_path)))

    ai_df["_title_key"] = ai_df["title"].apply(normalise_title)
    lf_df["_title_key"] = lf_df["title"].apply(normalise_title)

    results = []
    for _, row in lf_df.iterrows():
        tk = normalise_title(row["title"])
        abstract = "" if pd.isna(row.get("abstract", "")) else str(row.get("abstract", "")).strip()
        match = ai_df[ai_df["_title_key"] == tk]
        if match.empty:
            results.append({"title": row["title"], "abstract": abstract,
                            "found": False, "ai_decision": "not_found"})
        else:
            ai_dec = normalise_decision(match.iloc[0]["screening_decision"])
            captured = ai_dec in ("maybe", "include")
            results.append({
                "title": row["title"],
                "abstract": abstract,
                "found": True,
                "ai_decision": ai_dec,
                "captured": captured,
            })

    df_results = pd.DataFrame(results)
    n_total = len(df_results)
    n_found = int(df_results["found"].sum())
    n_not_found = n_total - n_found

    found_df = df_results[df_results["found"]].copy()
    if not found_df.empty:
        found_df["captured"] = found_df["captured"].astype(bool)
    n_captured = int(found_df["captured"].sum()) if not found_df.empty else 0
    n_missed = int((~found_df["captured"]).sum()) if not found_df.empty else 0
    capture_rate = n_captured / n_found if n_found > 0 else float("nan")
    miss_rate = n_missed / n_found if n_found > 0 else float("nan")

    missed_titles = found_df[~found_df["captured"]]["title"].tolist() if not found_df.empty else []
    missed_articles = (
        found_df[~found_df["captured"]][["title", "abstract"]].to_dict("records")
        if not found_df.empty else []
    )

    return {
        "n_listfinal": n_total,
        "n_found": n_found,
        "n_not_found": n_not_found,
        "n_captured": n_captured,
        "n_missed": n_missed,
        "capture_rate": capture_rate,
        "miss_rate": miss_rate,
        "missed_titles": missed_titles,
        "missed_articles": missed_articles,
    }


# ── Test-retest (reproducibility) ────────────────────────────────────

def run_test_retest(path_t1: Path, path_t2: Path):
    """Compare test 1 vs test 2 of the same model.  Returns dict with results."""
    t1_df = normalise_columns(load_file(str(path_t1)))
    t2_df = normalise_columns(load_file(str(path_t2)))

    for df in (t1_df, t2_df):
        df["_title_key"] = df["title"].apply(normalise_title)
        df["_occ"] = df.groupby("_title_key").cumcount().astype(str)
        df["_merge_key"] = df["_title_key"] + "__" + df["_occ"]

    merged = pd.merge(
        t1_df[["title", "_merge_key", "screening_decision"]],
        t2_df[["_merge_key", "screening_decision"]],
        on="_merge_key",
        how="inner",
        suffixes=("_t1", "_t2"),
    )

    n_total = len(merged)

    merged["t1_bin"] = merged["screening_decision_t1"].apply(binarise_decision)
    merged["t2_bin"] = merged["screening_decision_t2"].apply(binarise_decision)

    merged["t1_orig"] = merged["screening_decision_t1"].apply(normalise_decision)
    merged["t2_orig"] = merged["screening_decision_t2"].apply(normalise_decision)
    exact_match = int((merged["t1_orig"] == merged["t2_orig"]).sum())
    exact_pct = exact_match / n_total if n_total > 0 else float("nan")

    binary_match = int((merged["t1_bin"] == merged["t2_bin"]).sum())
    binary_pct = binary_match / n_total if n_total > 0 else float("nan")

    tp = int(((merged["t1_bin"] == "maybe")   & (merged["t2_bin"] == "maybe")).sum())
    fp = int(((merged["t1_bin"] == "maybe")   & (merged["t2_bin"] == "exclude")).sum())
    fn = int(((merged["t1_bin"] == "exclude") & (merged["t2_bin"] == "maybe")).sum())
    tn = int(((merged["t1_bin"] == "exclude") & (merged["t2_bin"] == "exclude")).sum())
    k, se, ci_lo, ci_hi, interp = calc_kappa(tp, fp, fn, tn)

    disc = merged[merged["t1_bin"] != merged["t2_bin"]]
    disc_titles = disc["title"].tolist()

    return {
        "n_total": n_total,
        "exact_match": exact_match,
        "exact_pct": exact_pct,
        "binary_match": binary_match,
        "binary_pct": binary_pct,
        "tp": tp, "fp": fp, "fn": fn, "tn": tn,
        "kappa": k, "kappa_se": se,
        "kappa_ci_lo": ci_lo, "kappa_ci_hi": ci_hi,
        "kappa_interp": interp,
        "n_discordant": len(disc),
        "disc_titles": disc_titles[:50],
    }
