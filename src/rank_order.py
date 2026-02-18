#!/usr/bin/env python3
"""Rank-order MAcc CORE courses by perceived benefit from exit survey data."""

from __future__ import annotations

import argparse
import string
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


DATA_START_ROW_IDX = 3  # Excel row 4 (0-based)

COURSE_COLUMN_MAP = {
    "L": "ACC 6060 Professionalism and Leadership",
    "M": "ACC 6300 Data Analytics",
    "N": "ACC 6400 Advanced Tax Business Entities",
    "O": "ACC 6510 Financial Audit",
    "P": "ACC 6540 Professional Ethics",
    "Q": "ACC 6560 Financial Theory & Research I",
    "R": "ACC 6350 Management Control Systems",
    "S": "ACC 6600 Business Law for Accountants",
}


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


def excel_col_to_index(col_letters: str) -> int:
    col_letters = col_letters.strip().upper()
    if not col_letters or any(ch not in string.ascii_uppercase for ch in col_letters):
        raise ValueError(f"Invalid Excel column letters: {col_letters!r}")

    idx = 0
    for ch in col_letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def load_rank_data(input_path: Path) -> tuple[pd.DataFrame, int, int]:
    if not input_path.exists():
        raise FileNotFoundError(f"Input file does not exist: {input_path}")

    df = pd.read_excel(input_path, engine="openpyxl", header=None)

    rank_col_letters = list(COURSE_COLUMN_MAP.keys())
    rank_col_indices = [excel_col_to_index(letter) for letter in rank_col_letters]
    course_names = [COURSE_COLUMN_MAP[letter] for letter in rank_col_letters]

    # Student responses start on Excel row 4 (index 3).
    rank_df = df.iloc[DATA_START_ROW_IDX:, rank_col_indices].copy()
    rank_df.columns = course_names

    # Convert all values to numeric ranks (coerce invalid to NaN).
    rank_df = rank_df.apply(pd.to_numeric, errors="coerce")

    total_rows_after_row4 = len(rank_df)
    rank_df = rank_df.dropna(how="all")
    rows_dropped_all_blank = total_rows_after_row4 - len(rank_df)

    return rank_df, total_rows_after_row4, rows_dropped_all_blank


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
        "Mean rank computed using available (non-missing) responses; n varies by course."
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

    rank_df, total_rows_after_row4, rows_dropped_all_blank = load_rank_data(input_path)

    print(f"Audit: total rows after row 4 = {total_rows_after_row4}")
    print(f"Audit: rows dropped as all-blank = {rows_dropped_all_blank}")
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
