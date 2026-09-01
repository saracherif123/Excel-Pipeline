import pandas as pd
import os
import glob
import re
from datetime import datetime
from urllib.parse import unquote
import argparse


FOLDER_PATH = "/Users/sarasaad/Documents/Data Processing/Data/P2F/P2F WH"
OUTPUT_PATH = "/Users/sarasaad/Documents/Data Processing/Data/P2F/P2F WH/processed"


def parse_version(name: str) -> str:
    match = re.search(r'_V([0-9]+(?:\.[0-9]+)?)', name, re.IGNORECASE)
    if match:
        v = match.group(1)
        if '.' not in v:
            v += '.0'
        return f"V{v}"
    return "V1.0"


def strip_version(name: str) -> str:
    return re.sub(r'_V[0-9]+(?:\.[0-9]+)?', '', name, flags=re.IGNORECASE).strip()


def strip_human_date(name: str) -> str:
    """Remove human-readable dates from filename, keeping only the ISO timestamp if present."""
    result = re.sub(r'_\d{2}-[A-Za-z]{3}-\d{4}', '', name)    # _07-Mar-2026
    result = re.sub(r'_[A-Za-z]{3}_\d{1,2}_\d{4}', '', result) # _Mar_7_2026 or _Feb_16_2026
    result = re.sub(r'_\d{4}-\d{2}-\d{2}(?!T)', '', result)    # _2026-03-07 (not ISO timestamp)
    return result


def parse_review_date(date_val):
    """Return (month_str, day_str, year_str, iso_str) from a date value."""
    try:
        if pd.isna(date_val):
            return "", "", "", ""
    except Exception:
        pass
    if isinstance(date_val, str):
        date_val = pd.to_datetime(date_val, errors='coerce')
    try:
        if pd.isna(date_val):
            return "", "", "", ""
    except Exception:
        pass
    month = date_val.strftime("%b")
    day   = str(date_val.day)
    year  = date_val.strftime("%Y")
    iso   = date_val.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    return month, day, year, iso


def build_deliverable_version_name(base_name: str, review_date) -> str:
    """
    Format: {clean_name}_{Version}_{Month}_{Day}_{Year}_{ISO}
    e.g. FSD_DA0005814_Q&FS_Flexible Inspection_V1.0_Mar_10_2026_2026-03-10T00:00:00.000Z
    """
    version               = parse_version(base_name)
    clean_name            = strip_version(base_name)
    clean_name            = strip_human_date(clean_name)  # remove any existing date from filename
    month, day, year, iso = parse_review_date(review_date)
    parts = [clean_name, version, month, day, year, iso]
    return "_".join(p for p in parts if p)


def clean_criteria_id(val: str) -> str:
    match = re.match(r'^([A-Za-z]+-C[0-9]+)', val.strip())
    return match.group(1) if match else val.strip()


def process_file(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath)
    df.columns = [c.strip() for c in df.columns]

    # Always drop the last row
    if not df.empty:
        df = df.iloc[:-1].reset_index(drop=True)
        print(f"  -> Dropped last row")

    raw_filename     = os.path.splitext(os.path.basename(filepath))[0]
    decoded_filename = unquote(raw_filename)
    clean_filename   = strip_human_date(decoded_filename)  # remove any date baked into filename

    # Auto-detect criteria and review date columns
    criteria_col    = next((c for c in df.columns if 'criteria' in c.lower()), None)
    review_date_col = next((c for c in df.columns if 'review date' in c.lower()), None)

    if criteria_col:
        print(f"  -> Criteria column found: '{criteria_col}'")
    else:
        print(f"  -> WARNING: No criteria column found")

    # --- Column A: Name = clean_filename + ISO date + CriteriaID (per row) ---
    def make_name(row):
        criteria_id = ""
        if criteria_col and pd.notna(row.get(criteria_col)):
            criteria_id = clean_criteria_id(str(row[criteria_col]))

        rd = row[review_date_col] if review_date_col else None
        _, _, _, iso_date = parse_review_date(rd)

        parts = [clean_filename, iso_date, criteria_id]
        return "_".join(p for p in parts if p)

    df.insert(0, 'Name', df.apply(make_name, axis=1))

    # --- Column G: Deliverable VersionName = clean_name + version + date (Month_Day_Year_ISO) ---
    def make_dvn(row):
        rd = row[review_date_col] if review_date_col else None
        return build_deliverable_version_name(decoded_filename, rd)

    dvn_series = df.apply(make_dvn, axis=1)

    if 'Deliverable VersionName' in df.columns:
        df.drop(columns=['Deliverable VersionName'], inplace=True)

    df['Deliverable VersionName'] = dvn_series
    cols = list(df.columns)
    cols.remove('Deliverable VersionName')
    cols.insert(min(6, len(cols)), 'Deliverable VersionName')
    df = df[cols]

    return df


def main():
    parser = argparse.ArgumentParser(description="Process Excel files (P2F naming + DVN generation).")
    parser.add_argument("--input", default=FOLDER_PATH, help="Folder containing .xlsx/.xls files")
    parser.add_argument("--output", default=OUTPUT_PATH, help="Folder to write processed files")
    args = parser.parse_args()

    os.makedirs(args.output, exist_ok=True)

    excel_files = glob.glob(os.path.join(args.input, "*.xlsx")) + glob.glob(os.path.join(args.input, "*.xls"))

    if not excel_files:
        print(f"No Excel files found in: {args.input}")
        return

    print(f"Found {len(excel_files)} file(s) to process.\n")

    for filepath in excel_files:
        filename = os.path.basename(filepath)
        print(f"Processing: {filename}")
        try:
            df = process_file(filepath)

            name_no_ext, ext = os.path.splitext(filename)
            out_path = os.path.join(args.output, f"{name_no_ext}{ext}")
            df.to_excel(out_path, index=False)

            print(f"  -> Saved: {out_path}")
            print(f"  -> Rows: {len(df)}")
            for i, row in df.head(2).iterrows():
                print(f"  -> Row {i+1} Name:                     {row['Name']}")
                if 'Deliverable VersionName' in df.columns:
                    print(f"  -> Row {i+1} Deliverable VersionName: {row['Deliverable VersionName']}")
        except Exception as e:
            print(f"  ERROR: {e}")
        print()

    print("Done.")


if __name__ == "__main__":
    main()