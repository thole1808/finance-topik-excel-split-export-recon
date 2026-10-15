"""Backward-compatible command-line entry point for Recon Builder."""

from __future__ import annotations

import argparse
from pathlib import Path

from src.recon_exporter import (
    LB_COL_MAP,
    RB_COL_MAP,
    build_recon_workbook,
    read_raw_excel,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Create a reconciliation workbook from LB.xlsx and RB.xlsx."
    )
    parser.add_argument("lb", type=Path, help="Input Loket Bayar workbook")
    parser.add_argument("rb", type=Path, help="Input Rajabiller workbook")
    parser.add_argument(
        "-o", "--output", type=Path, default=Path("recon-output.xlsx"),
        help="Output workbook path",
    )
    args = parser.parse_args()

    lb_df = read_raw_excel(args.lb, col_map=LB_COL_MAP)
    rb_df = read_raw_excel(
        args.rb, col_map=RB_COL_MAP, drop_col_non_numeric="no"
    )
    output = build_recon_workbook(lb_df, rb_df, args.output)
    print(f"Recon selesai: {output}")


if __name__ == "__main__":
    main()
