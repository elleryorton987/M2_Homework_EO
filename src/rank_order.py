#!/usr/bin/env python3
"""Rank-order MAcc CORE courses by perceived benefit from exit survey data."""

from __future__ import annotations

import argparse
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Rank-order courses from the 2024 MAcc Exit Survey rank-order question."
    )
    parser.add_argument("--input", required=True, help="Path to input Excel file")
    parser.add_argument("--outdir", required=True, help="Output directory")
    return parser.parse_args()


def load_rank_data(input_path: Path) -> pd.DataFrame:
    if not input_path.exists():
        raise FileNotFoundError(f"Input file does not exist: {input_path}")

    return pd.read_excel(input_path, header=0)


def summarize_rankings(rank_df: pd.DataFrame) -> pd.DataFrame:
    row_count = len(rank_df)
    summary = pd.DataFrame(
        {
            "course_name": rank_df.columns,
            "n": [row_count for _ in rank_df.columns],
            "mean_rank": [float(rank_df[c].mean()) for c in rank_df.columns],
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
    ax.set_xlabel("mean_rank")
    ax.set_ylabel("Course Name")
    ax.set_title("2024 MAcc Exit Survey: CORE Course Benefit Ranking", pad=14)
    fig.text(
        0.5,
        0.93,
        "Mean rank (1 = most beneficial). Lower is better.",
        ha="center",
        va="center",
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

    summary = summarize_rankings(rank_df)

    csv_path = outdir / "course_rankings.csv"
    png_path = outdir / "rank_order.png"

    summary.to_csv(csv_path, index=False)
    save_chart(summary, png_path)

    print(f"Saved: {csv_path}")
    print(f"Saved: {png_path}")


if __name__ == "__main__":
    main()
