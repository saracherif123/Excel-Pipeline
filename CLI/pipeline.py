"""
Run the Excel pipeline as a module.

Typical usage:
  python3 -m src.pipeline --input-folder "/path/to/raw" --processed-folder "/path/to/processed" --combined-file "/path/to/combined.xlsx"
"""

from __future__ import annotations

import argparse
import os

from src.combiner import combine_excels
from src.data_q2 import clean_combined_file


def main():
    parser = argparse.ArgumentParser(description="End-to-end Excel pipeline runner (combine + clean).")
    parser.add_argument("--combine-input", required=True, help="Folder containing Excel files to combine")
    parser.add_argument("--combined-output", required=True, help="Output combined .xlsx file")
    parser.add_argument("--clean-output", required=False, help="Output cleaned .xlsx file (optional)")
    parser.add_argument("--no-summary", action="store_true", help="Do not write the Summary sheet during combine")
    args = parser.parse_args()

    os.makedirs(os.path.dirname(args.combined_output) or ".", exist_ok=True)
    combine_excels(args.combine_input, args.combined_output, add_summary_sheet=not args.no_summary)

    if args.clean_output:
        os.makedirs(os.path.dirname(args.clean_output) or ".", exist_ok=True)
        clean_combined_file(args.combined_output, args.clean_output)


if __name__ == "__main__":
    main()

