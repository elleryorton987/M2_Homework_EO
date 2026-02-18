#!/usr/bin/env python3
"""Rank-order MAcc CORE courses by perceived benefit from exit survey data."""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


COL_START = "L"
COL_END = "S"
LABEL_ROW_IDX = 1  # Excel row 2
RESPONSES_START_IDX = 3  # Excel row 4

REFLECTION_TEMPLATE = """- What changed from Project 1 to this workflow?
- Where is the control now?
- What would you do next if you had one more week?
- Identify one accounting application of this workflow from another class you have taken or are taking. Be specific.
"""


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Rank-order courses from the 2024 MAcc Exit Survey rank-order question."
    )
    parser.add_argument("--input", required=True, help="Path to input Excel file")
    parser.add_argument("--outdir", required=True, help="Output directory")
    return parser.parse_args()


def parse_course_name(label: object) -> str:
    label_text = "" if pd.isna(label) else str(label).strip()
    if " - " in label_text:
        return label_text.split(" - ")[-1].strip()
    return label_text


def load_rank_data(input_path: Path) -> pd.DataFrame:
    if not input_path.exists():
        raise FileNotFoundError(f"Input file does not exist: {input_path}")

    raw = pd.read_excel(
        input_path,
        engine="openpyxl",
        header=None,
        usecols=f"{COL_START}:{COL_END}",
    )

    course_names = [parse_course_name(value) for value in raw.iloc[LABEL_ROW_IDX].tolist()]

    responses = raw.iloc[RESPONSES_START_IDX:].copy()
    responses.columns = course_names

    for col in responses.columns:
        responses[col] = pd.to_numeric(responses[col], errors="coerce")

    return responses


def summarize_rankings(rank_df: pd.DataFrame) -> pd.DataFrame:
    summary = pd.DataFrame(
        {
            "course_name": rank_df.columns,
            "n": [int(rank_df[c].notna().sum()) for c in rank_df.columns],
            "mean_rank": [float(rank_df[c].mean(skipna=True)) for c in rank_df.columns],
        }
    )
    summary = summary.sort_values(by="mean_rank", ascending=True, kind="mergesort").reset_index(
        drop=True
    )
    summary["final_rank"] = summary.index + 1
    return summary


def save_chart(summary: pd.DataFrame, outpath: Path) -> None:
    plot_df = summary.sort_values("final_rank", ascending=True)

    fig, ax = plt.subplots(figsize=(10, 6))
    ax.barh(plot_df["course_name"], plot_df["mean_rank"], color="#4C78A8")
    ax.invert_yaxis()  # rank 1 at top
    ax.set_xlabel("Mean Rank")
    ax.set_ylabel("Course Name")
    ax.set_title(
        "2024 MAcc Exit Survey: CORE Course Benefit Ranking\n"
        "Mean rank (1 = most beneficial). Lower is better."
    )
    ax.grid(axis="x", linestyle="--", alpha=0.4)
    fig.tight_layout()
    fig.savefig(outpath, dpi=200)
    plt.close(fig)


def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    rank_df = load_rank_data(input_path)

    print(f"Audit: total rows read = {len(rank_df)}")
    print(f"Audit: extracted course_name list = {list(rank_df.columns)}")
    for course in rank_df.columns:
        print(f"Audit: {course} n = {int(rank_df[course].notna().sum())}")

    summary = summarize_rankings(rank_df)

    csv_path = outdir / "course_rankings.csv"
    png_path = outdir / "rank_order.png"
    reflection_path = outdir / "reflection.md"

    summary.to_csv(csv_path, index=False)
    save_chart(summary, png_path)
    reflection_path.write_text(REFLECTION_TEMPLATE, encoding="utf-8")

    print(f"Saved: {csv_path}")
    print(f"Saved: {png_path}")
    print(f"Saved: {reflection_path}")


if __name__ == "__main__":
    main()
