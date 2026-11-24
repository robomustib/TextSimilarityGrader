"""
TextSimilarityGrader (https://github.com/robomustib/TextSimilarityGrader/)
Copyright (c) 2026 Mustafa Bilgin
Licensed under Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)
Add-on: Sensitivity analysis of the fuzzy tolerance (DE/TR)

Re-grades a set of transcripts with the unchanged grading logic of
evaluate_BLCK_DE_TR.py across a range of tolerance values and compares
each run with a human reference rating (gold standard).

For every tolerance the script reports
  - false negatives (reference 1, automated 0)
  - false positives (reference 0, automated 1)
  - total number of deviations
  - raw agreement and Cohen's kappa
separately for each language and for all items combined.

Tolerance = 1 - FUZZY_THRESHOLD (e.g. tolerance 30 % = threshold 0.70).
The stricter rule for targets with at most three characters (minimum
similarity 90 %) is part of the grading logic and is not varied.

Usage
    python sensitivity_analysis.py <input.xlsx> [options]

Example (column names of the ZDKM study data)
    python sensitivity_analysis.py F2_long.xlsx \
        --col-transcript ai_transcript_full_response \
        --col-target target_words --col-forbidden forbidden \
        --col-gold gold_point --col-language test_lang
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

# The grading functions are imported from the published evaluation script,
# so the sensitivity analysis always uses exactly the same rules.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import evaluate_BLCK_DE_TR as grader  # noqa: E402

LANGUAGE_MAP = {"DE": "DE", "GE": "DE", "GER": "DE", "DEU": "DE",
                "TR": "TR", "TUR": "TR"}


def grade_item(transcript, target, forbidden, language, threshold):
    """Grades one item with the logic of evaluate_BLCK_DE_TR.py."""
    grader.EVALUATION_LANGUAGE = language
    grader.FUZZY_THRESHOLD = threshold

    if pd.isna(transcript) or not str(transcript).strip():
        return 0
    if pd.isna(target) or not str(target).strip():
        return 0

    match_word, _sim, points, matched_target = grader.find_best_match(
        str(target), str(transcript), grader.SCORING_MODE)

    is_forbidden, _ = grader.check_forbidden(
        forbidden, str(transcript),
        match_word if points > 0 else None,
        matched_target if points > 0 else None)

    return 0 if is_forbidden else points


def agreement(reference, automated):
    """Raw agreement (%) and Cohen's kappa for two binary ratings."""
    ref = pd.Series(reference).astype(int).to_numpy()
    aut = pd.Series(automated).astype(int).to_numpy()
    n = len(ref)
    if n == 0:
        return float("nan"), float("nan")
    po = (ref == aut).mean()
    p_ref, p_aut = ref.mean(), aut.mean()
    pe = p_ref * p_aut + (1 - p_ref) * (1 - p_aut)
    kappa = (po - pe) / (1 - pe) if pe < 1 else float("nan")
    return round(100 * po, 1), round(kappa, 3)


def summarise(df, tolerance, label):
    ref, aut = df["_gold"], df["_points"]
    fn = int(((ref == 1) & (aut == 0)).sum())
    fp = int(((ref == 0) & (aut == 1)).sum())
    po, kappa = agreement(ref, aut)
    return {"Tolerance (%)": tolerance, "Language": label, "n": len(df),
            "False negatives": fn, "False positives": fp,
            "Deviations total": fn + fp,
            "Agreement (%)": po, "Cohen's kappa": kappa}


def main():
    parser = argparse.ArgumentParser(
        description="Sensitivity analysis of the fuzzy tolerance.")
    parser.add_argument("input", help="Excel file with one row per item")
    parser.add_argument("--col-transcript", default="Transcript")
    parser.add_argument("--col-target", default="Target")
    parser.add_argument("--col-forbidden", default="Forbidden")
    parser.add_argument("--col-gold", default="Gold_Point")
    parser.add_argument("--col-language", default="Language")
    parser.add_argument("--start", type=int, default=0,
                        help="lowest tolerance in percent (default 0)")
    parser.add_argument("--stop", type=int, default=50,
                        help="highest tolerance in percent (default 50)")
    parser.add_argument("--step", type=int, default=5,
                        help="step width in percent (default 5)")
    parser.add_argument("--output", default="Sensitivity_Results.xlsx")
    args = parser.parse_args()

    df = pd.read_excel(args.input)
    needed = [args.col_transcript, args.col_target, args.col_forbidden,
              args.col_gold, args.col_language]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        sys.exit(f"Missing columns: {', '.join(missing)}\n"
                 f"Available columns: {', '.join(map(str, df.columns))}")

    df["_lang"] = (df[args.col_language].astype(str).str.strip().str.upper()
                   .map(LANGUAGE_MAP))
    if df["_lang"].isna().any():
        bad = sorted(df.loc[df["_lang"].isna(), args.col_language]
                     .astype(str).unique())
        sys.exit(f"Unknown language codes: {', '.join(bad)} (use DE or TR)")
    df["_gold"] = df[args.col_gold].astype(int)

    rows, item_points = [], {}
    for tolerance in range(args.start, args.stop + 1, args.step):
        threshold = round(1 - tolerance / 100, 4)
        df["_points"] = [
            grade_item(t, tg, fb, lg, threshold)
            for t, tg, fb, lg in zip(df[args.col_transcript],
                                     df[args.col_target],
                                     df[args.col_forbidden],
                                     df["_lang"])]
        item_points[f"points_tol_{tolerance}"] = df["_points"].copy()
        for lang in sorted(df["_lang"].unique()):
            rows.append(summarise(df[df["_lang"] == lang], tolerance, lang))
        rows.append(summarise(df, tolerance, "ALL"))

    summary = pd.DataFrame(rows)
    overall = summary[summary["Language"] == "ALL"]
    best = overall.loc[overall["Deviations total"].idxmin()]

    print("=" * 60)
    print(" SENSITIVITY ANALYSIS OF THE FUZZY TOLERANCE")
    print("=" * 60)
    print(overall[["Tolerance (%)", "False negatives", "False positives",
                   "Deviations total", "Cohen's kappa"]]
          .to_string(index=False))
    print("-" * 60)
    print(f" Minimum of deviations at a tolerance of "
          f"{int(best['Tolerance (%)'])} % "
          f"({int(best['Deviations total'])} deviations)")

    items = df.drop(columns=["_points"]).copy()
    for name, values in item_points.items():
        items[name] = values.values
    items = items.drop(columns=["_lang", "_gold"])

    with pd.ExcelWriter(args.output) as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        items.to_excel(writer, sheet_name="Items", index=False)
    print(f"\n Results saved as: {args.output}")


if __name__ == "__main__":
    main()
