"""
Combine multiple Excel files into one consolidated workbook.

Can be used as a script or imported as a function.
"""

from __future__ import annotations

import argparse
import glob
import os
from typing import Iterable

import pandas as pd


DEFAULT_INPUT_PATH = "/Users/sarasaad/Documents/Data Processing/Data/P2F/ALL"
DEFAULT_OUTPUT_FILE = "/Users/sarasaad/Documents/Data Processing/Data/P2F/ALL/combined_output.xlsx"

_TOTAL_ROW_PAT = r"(?i)^\s*(grand\s+total|total)\s*$"


def _drop_trailing_total_row(df: pd.DataFrame) -> tuple[pd.DataFrame, bool]:
    """
    If the last row looks like a totals/footer row, drop it.

    Heuristic: in the last row, any string cell equals 'total'/'grand total'
    (case-insensitive, whitespace tolerant).
    """
    if df.empty:
        return df, False

    last = df.iloc[-1]
    for v in last.values.tolist():
        if isinstance(v, str) and v is not None:
            s = v.strip()
            if not s:
                continue
            if pd.Series([s]).astype("string").str.match(_TOTAL_ROW_PAT).iloc[0]:
                return df.iloc[:-1].reset_index(drop=True), True

    # Heuristic 2: if there is a Score column and last-row score is not in {0, 0.5, 1},
    # it's likely a totals/footer row.
    score_col = None
    for c in df.columns:
        if str(c).strip().lower() == "score":
            score_col = c
            break
    if score_col is not None:
        v = df.iloc[-1][score_col]
        # Blank score on the last row usually indicates a footer row.
        if pd.isna(v) or (isinstance(v, str) and not v.strip()):
            return df.iloc[:-1].reset_index(drop=True), True

        # If it's a string like "Total" or similar, drop.
        if isinstance(v, str):
            if pd.Series([v.strip()]).astype("string").str.match(_TOTAL_ROW_PAT).iloc[0]:
                return df.iloc[:-1].reset_index(drop=True), True

        # If it's numeric but outside the allowed set, drop.
        vv = pd.to_numeric(pd.Series([v]), errors="coerce").iloc[0]
        if pd.notna(vv) and float(vv) not in (0.0, 0.5, 1.0):
            return df.iloc[:-1].reset_index(drop=True), True
    return df, False


def _list_excel_files(input_path: str) -> list[str]:
    files = glob.glob(os.path.join(input_path, "*.xlsx")) + glob.glob(os.path.join(input_path, "*.xls"))
    # Exclude temp Excel files and an existing combined output
    out = []
    for fp in files:
        base = os.path.basename(fp)
        if base.startswith("~$"):
            continue
        if base.lower() == os.path.basename(DEFAULT_OUTPUT_FILE).lower():
            continue
        out.append(fp)
    return sorted(out)


def align_union_schema(dfs: Iterable[pd.DataFrame]) -> tuple[list[pd.DataFrame], list[str]]:
    dfs = list(dfs)
    if not dfs:
        return [], []

    all_cols: list[str] = []
    seen: set[str] = set()
    for df in dfs:
        for c in df.columns:
            if c not in seen:
                seen.add(c)
                all_cols.append(c)

    meta_first = [c for c in ["_source_file", "_source_sheet"] if c in seen]
    rest = [c for c in all_cols if c not in set(meta_first)]
    ordered = meta_first + rest

    return [df.reindex(columns=ordered) for df in dfs], ordered


def combine_folder_to_frames(input_path: str) -> dict:
    """
    Load all Excel files in a folder and return:
      - combined_df: concatenated dataframe (union schema)
      - summary_df: rows per source file (if _source_file present)
      - meta: counts
    Does not write any files (useful for web apps / in-memory export).
    """
    excel_files = _list_excel_files(input_path)
    if not excel_files:
        raise FileNotFoundError(f"No Excel files (.xlsx/.xls) found in: {input_path}")

    dfs: list[pd.DataFrame] = []
    loaded = 0
    skipped = 0
    total_footer_rows_dropped = 0
    for filepath in excel_files:
        filename = os.path.basename(filepath)
        try:
            df = pd.read_excel(filepath)
            df["_source_file"] = filename
            df, dropped = _drop_trailing_total_row(df)
            if dropped:
                total_footer_rows_dropped += 1
            if df.dropna(how="all").empty:
                skipped += 1
                continue
            dfs.append(df)
            loaded += 1
        except Exception:
            skipped += 1

    if not dfs:
        raise RuntimeError("All files failed to load; nothing to combine.")

    aligned, ordered_cols = align_union_schema(dfs)
    combined_df = pd.concat(aligned, ignore_index=True, sort=False, copy=False)
    summary_df = None
    if "_source_file" in combined_df.columns:
        summary_df = (
            combined_df.groupby("_source_file", dropna=False)
            .size()
            .reset_index(name="rows")
            .sort_values("rows", ascending=False)
        )

    return {
        "combined_df": combined_df,
        "summary_df": summary_df,
        "meta": {
            "files_found": len(excel_files),
            "files_loaded": loaded,
            "files_skipped": skipped,
            "total_footer_rows_dropped": total_footer_rows_dropped,
            "total_rows": len(combined_df),
            "total_cols": len(ordered_cols),
        },
    }


def combine_excels(input_path: str, output_file: str, add_summary_sheet: bool = True) -> dict:
    excel_files = _list_excel_files(input_path)
    if not excel_files:
        raise FileNotFoundError(f"No Excel files (.xlsx/.xls) found in: {input_path}")

    print(f"Found {len(excel_files)} file(s) to combine.\n")

    dfs: list[pd.DataFrame] = []
    loaded = 0
    skipped = 0
    total_footer_rows_dropped = 0
    for filepath in excel_files:
        filename = os.path.basename(filepath)
        print(f"  Reading: {filename}")
        try:
            df = pd.read_excel(filepath)
            df["_source_file"] = filename
            df, dropped = _drop_trailing_total_row(df)
            if dropped:
                total_footer_rows_dropped += 1
            # Skip files that are empty / effectively empty.
            if df.dropna(how="all").empty:
                skipped += 1
                print("    -> SKIPPED: empty sheet")
                continue

            dfs.append(df)
            loaded += 1
            print(f"    -> {len(df)} rows, {len(df.columns)} columns")
        except Exception as e:
            skipped += 1
            print(f"    -> SKIPPED: {e}")

    if not dfs:
        raise RuntimeError("All files failed to load; nothing to combine.")

    aligned, ordered_cols = align_union_schema(dfs)
    combined = pd.concat(aligned, ignore_index=True, sort=False, copy=False)

    os.makedirs(os.path.dirname(output_file) or ".", exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as w:
        combined.to_excel(w, index=False, sheet_name="Combined")
        if add_summary_sheet and "_source_file" in combined.columns:
            summary = (
                combined.groupby("_source_file", dropna=False)
                .size()
                .reset_index(name="rows")
                .sort_values("rows", ascending=False)
            )
            summary.to_excel(w, index=False, sheet_name="Summary")

    print(f"\nTotal rows combined: {len(combined)}")
    print(f"Saved combined file to: {output_file}")

    return {
        "files_found": len(excel_files),
        "files_loaded": loaded,
        "files_skipped": skipped,
        "total_footer_rows_dropped": total_footer_rows_dropped,
        "total_rows": len(combined),
        "total_cols": len(ordered_cols),
        "output_file": output_file,
    }


def main():
    parser = argparse.ArgumentParser(description="Combine Excel files into a consolidated workbook.")
    parser.add_argument("--input", default=DEFAULT_INPUT_PATH, help="Folder containing .xlsx/.xls files")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_FILE, help="Output .xlsx path")
    parser.add_argument("--no-summary", action="store_true", help="Do not write the Summary sheet")
    args = parser.parse_args()

    combine_excels(args.input, args.output, add_summary_sheet=not args.no_summary)


if __name__ == "__main__":
    main()