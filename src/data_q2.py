import re
import pandas as pd
from datetime import datetime

import argparse
from typing import Optional

DEFAULT_INPUT_FILE = "/Users/sarasaad/Documents/Data Processing/Data/P2F/ALL/combined_output.xlsx"
DEFAULT_OUTPUT_FILE = "/Users/sarasaad/Documents/Data Processing/Data/P2F/ALL/combined_output_fin.xlsx"

def standardize_date(date_val):
    if pd.isna(date_val):
        return None
    date_str = (
        str(date_val).strip()
        .replace("‑", "-")  # non-breaking hyphen
        .replace("–", "-")  # en dash
        .replace("—", "-")  # em dash
    )
    formats = [
        "%d-%b-%Y",   # 07-Mar-2026
        "%Y-%m-%d",   # 2026-02-16
        "%d-%m-%Y",   # 16-02-2026
        "%m-%d-%Y",   # 02-16-2026
        "%d/%m/%Y",   # 16/02/2026
        "%m/%d/%Y",   # 02/16/2026
        "%Y/%m/%d",   # 2026/02/16
        "%B %d, %Y",  # February 16, 2026
        "%d %B %Y",   # 16 February 2026
    ]
    if isinstance(date_val, pd.Timestamp):
        return date_val.strftime("%Y-%m-%d")
    date_str = date_str.split(" ")[0]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    print(f"  WARNING: Could not parse date: '{date_str}'")
    return date_str

def remove_human_date(val):
    if pd.isna(val):
        return None
    result = str(val).strip()
    result = re.sub(r'_\d{2}-[A-Za-z]{3}-\d{4}', '', result)    # _07-Mar-2026
    result = re.sub(r'_[A-Za-z]{3}_\d{1,2}_\d{4}', '', result)  # _Mar_7_2026 or _Feb_16_2026
    result = re.sub(r'_\d{4}-\d{2}-\d{2}(?!T)', '', result)     # _2026-03-07 (not ISO timestamp)
    return result

def clean_criteria_id(criteria_val):
    if pd.isna(criteria_val):
        return None
    criteria_str = str(criteria_val).strip()
    # Accept patterns like:
    # - CRD-C001 ...
    # - QFS-C12 ...
    # - C001 ...
    # - CRD‑C001 ... (unicode hyphen)
    normalized = (
        criteria_str.replace("–", "-")
        .replace("—", "-")
        .replace("‑", "-")
    )
    match = re.match(r"^([A-Z]+-C\d+|[A-Z]+-[A-Z]\d+|C\d+)", normalized, flags=re.IGNORECASE)
    if match:
        return match.group(1).upper()
    print(f"  WARNING: Could not extract ID from: '{criteria_str}'")
    return normalized


def _coalesce_columns(df: pd.DataFrame, sources: list[str], target: str) -> pd.DataFrame:
    existing = [c for c in sources if c in df.columns]
    if not existing:
        return df
    ser = None
    for c in existing:
        s = df[c]
        ser = s if ser is None else ser.combine_first(s)
    df[target] = ser
    for c in existing:
        if c != target:
            df.drop(columns=[c], inplace=True)
    return df


def _normalize_crit_id_token(token) -> object:
    """
    Normalize a raw criteria-id token into the canonical short form ``TYPE-Cnnn``.

    Handles legacy bare-number IDs like ``FSD-01`` / ``FSD-1`` (no ``C`` prefix)
    by inserting the missing ``C`` and zero-padding to 3 digits, so they merge
    cleanly with the dominant ``FSD-C001`` style used by every other file.
    Existing IDs that already include the ``C`` are normalized to the same
    width (e.g. ``FSD-C1`` → ``FSD-C001``).
    """
    if token is None:
        return pd.NA
    try:
        if pd.isna(token):
            return pd.NA
    except (TypeError, ValueError):
        pass
    s = str(token).strip().upper()
    s = s.replace("–", "-").replace("—", "-").replace("‑", "-")
    m = re.match(r"^([A-Z]{2,})-C?0*(\d+)$", s)
    if not m:
        m2 = re.match(r"^C0*(\d+)$", s)
        if m2:
            return f"C{int(m2.group(1)):03d}"
        return s or pd.NA
    return f"{m.group(1)}-C{int(m.group(2)):03d}"


def _extract_criteria_id_from_text(s: pd.Series) -> pd.Series:
    if s is None:
        return s
    normalized = (
        s.astype("string")
        .str.replace("–", "-", regex=False)
        .str.replace("—", "-", regex=False)
        .str.replace("‑", "-", regex=False)
    )
    extracted = normalized.str.extract(r"(?i)\b([A-Z]+-C?\d+|C\d+)\b", expand=False)
    return extracted.apply(_normalize_crit_id_token).astype("string")


def fix_out_of_place_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Handle common consolidation issues:
    - Drop 'Unnamed:*' columns
    - Standardize source filename column
    - Coalesce duplicated comment columns into one
    - Extract/clean Criteria ID from text columns when missing
    - Standardize review_date if present
    """
    unnamed = [c for c in df.columns if str(c).lower().startswith("unnamed")]
    if unnamed:
        df = df.drop(columns=unnamed)

    if "source_file" in df.columns and "_source_file" not in df.columns:
        df = df.rename(columns={"source_file": "_source_file"})

    # Coalesce duplicate comment-like columns
    df = _coalesce_columns(
        df,
        sources=[
            "detailed_comments",
            "detailed_comments_2",
            "detailed_comments3",
            "detailed_comments_gaps_recommendations",
        ],
        target="detailed_comments",
    )
    df = _coalesce_columns(df, sources=["comments", "comments2"], target="comments")

    # Criteria ID: clean or extract if mostly missing
    if "criteria_id" in df.columns:
        df["criteria_id"] = df["criteria_id"].apply(clean_criteria_id)
    else:
        df["criteria_id"] = pd.NA

    if len(df):
        missing_ratio = float(df["criteria_id"].isna().mean())
    else:
        missing_ratio = 1.0

    if missing_ratio > 0.5:
        for col in ["criteria_id_deliverable_name", "criteria_id_number_deliverable_name", "name"]:
            if col in df.columns:
                extracted = _extract_criteria_id_from_text(df[col])
                df["criteria_id"] = df["criteria_id"].astype("string").combine_first(extracted)
        df["criteria_id"] = df["criteria_id"].apply(clean_criteria_id)

    if "review_date" in df.columns:
        df["review_date"] = df["review_date"].apply(standardize_date)

    preferred = [
        "_source_file",
        "name",
        "criteria_id",
        "criteria_name",
        "review_date",
        "status",
        "score",
        "deliverable_versionname",
        "deliverable",
        "timestamp",
        "comments",
        "detailed_comments",
    ]
    cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in set(preferred)]
    return df.reindex(columns=cols)


P2F_FINAL_COLUMNS = [
    "Name",
    "Criteria ID Number – Deliverable Name",
    "Criteria Name",
    "Review Date and Timestamp",
    "Status",
    "Score",
    "Deliverable VersionName",
    "Detailed comments",
]

QFS_FINAL_COLUMNS = [
    "Name",
    "Review Date",
    "Status",
    "Score",
    "Detailed Comments",
    "Deliverable VersionName",
    "Criteria Name",
    "Criteria ID",
    "Comment",
]


def _first_present(df: pd.DataFrame, names: list[str]) -> Optional[str]:
    for n in names:
        if n in df.columns:
            return n
    return None


def to_p2f_final_schema(df: pd.DataFrame) -> pd.DataFrame:
    """
    Reshape a consolidated dataframe (snake_case or Title Case) into the exact
    8-column schema used by `output/P2F_Final (3).xlsx`.

    Key rules derived from the reference file:
    - Criteria ID: short form TYPE-Cxxx (e.g. FSD-C001, CRD-C001, PDD-C001)
    - Review Date and Timestamp: kept as raw string (no datetime conversion)
    - Name: source Name passed through (remove_human_date only), criteria ID appended if not already present
    - DVN: source DVN passed through (remove_human_date), ISO timestamp appended if none present
    """
    df = df.copy()
    df = fix_out_of_place_columns(df)

    # --- Resolve source columns ---
    name_src = _first_present(df, ["Name", "name"])
    criteria_name_src = _first_present(df, ["Criteria Name", "criteria_name"])
    status_src = _first_present(df, ["Status", "status"])
    score_src = _first_present(df, ["Score", "score"])
    dvn_src = _first_present(df, ["Deliverable VersionName", "deliverable_versionname"])
    review_ts_src = _first_present(df, ["Review Date and Timestamp", "review_date_and_timestamp"])
    timestamp_src = _first_present(df, ["timestamp", "Timestamp"])
    review_date_src = _first_present(df, ["review_date", "Review Date"])
    source_file_src = _first_present(df, ["source_file", "_source_file"])
    detailed_src = _first_present(
        df,
        ["detailed_comments", "Detailed comments", "Detailed Comments",
         "Detailed comments (gaps & recommendations)"],
    )

    def _best_nonnull_col(candidates: list) -> Optional[str]:
        best, best_n = None, -1
        for c in candidates:
            if c in df.columns:
                n = int(df[c].notna().sum())
                if n > best_n:
                    best, best_n = c, n
        return best

    crit_num_src = _best_nonnull_col(
        ["Criteria ID Number – Deliverable Name", "criteria_id_number_deliverable_name"]
    )
    crit_alt_src = _best_nonnull_col(
        ["criteria_id_deliverable_name", "Criteria ID – Deliverable Name"]
    )

    # --- Criteria ID extraction ---
    # Produces the short form: FSD-C001, PDD-C001, CRD-C001 etc.
    def _extract_short_crit_id(s: pd.Series) -> pd.Series:
        s_str = (
            s.astype("string")
            .str.replace("–", "-", regex=False)
            .str.replace("—", "-", regex=False)
            .str.replace("‑", "-", regex=False)
        )
        # Pattern A: TYPE-C### at start (CRD style)
        # e.g. "CRD-C001 - description" -> "CRD-C001"
        p_a = s_str.str.extract(r"^([A-Z]{2,}-C\d+)", expand=False)
        # Pattern A2: TYPE-NN at start (bare number variant, no "C")
        # e.g. "FSD-01 - Daily Results Report" -> "FSD-C001"
        p_a2_raw = s_str.str.extract(r"^([A-Z]{2,}-\d{1,3})(?:\b|$|\s|-)", expand=False)
        p_a2 = p_a2_raw.where(p_a2_raw.isna(), p_a2_raw.apply(_normalize_crit_id_token).astype("string"))
        # Pattern B: TYPE_..._C### (FSD/PDD style: extract prefix + last plain _C-number)
        # e.g. "FSD_I005529A_..._C001" -> "FSD-C001"
        type_pre = s_str.str.extract(r"^([A-Z]{2,})_", expand=False)
        c_suf = s_str.str.extract(r"_(C\d+)$", expand=False)
        p_b = (type_pre + "-" + c_suf).where(type_pre.notna() & c_suf.notna())
        # Pattern C: ..._TYPE-C### or ..._TYPE_C### at end (new file format uses TYPE-Cxxx)
        # e.g. "FSD_..._FSD-C001" -> "FSD-C001",  "PDD_..._PDD-C001" -> "PDD-C001"
        type_suf = s_str.str.extract(r"_([A-Z]{2,})[_\-](C\d+)$", expand=True)
        p_c = (type_suf[0] + "-" + type_suf[1]).where(type_suf[0].notna() & type_suf[1].notna())
        result = p_a.combine_first(p_a2).combine_first(p_b).combine_first(p_c)
        # Normalize all separators between type and C-number to dash (CRD_C001 → CRD-C001)
        result = result.str.replace(r"([A-Z]{2,})[_\u2011\u2013](C\d+)", r"\1-\2", regex=True)
        # Final pass: zero-pad numeric portion so FSD-C1 / FSD-C01 → FSD-C001
        result = result.apply(_normalize_crit_id_token).astype("string")
        return result

    crit_id = pd.Series([pd.NA] * len(df), dtype="string", index=df.index)
    if crit_num_src:
        crit_id = crit_id.combine_first(_extract_short_crit_id(df[crit_num_src]))
    if crit_alt_src:
        crit_id = crit_id.combine_first(_extract_short_crit_id(df[crit_alt_src]))

    # For bare Cxxx (PDD files with only "Criteria ID" = "C001"):
    # infer type prefix from source Name, source_file, or deliverable column.
    bare_c_mask = crit_id.isna()
    if bare_c_mask.any() and "criteria_id" in df.columns:
        bare_c = df["criteria_id"].astype("string").str.extract(r"^(C\d+)$", expand=False)
        name_for_type = (
            df[name_src].astype("string") if name_src
            else pd.Series([""] * len(df), index=df.index)
        )
        type_from_name = (
            name_for_type.str.extract(r"(?i)(?:^|_)(FSD|CRD|PDD)(?:_|$)", expand=False)
            .astype("string").str.upper()
        )
        # Fall back to source_file when name column is absent or doesn't contain the type
        if source_file_src:
            sf_type = (
                df[source_file_src].astype("string")
                .str.extract(r"(?i)(?:^|_)(FSD|CRD|PDD)(?:_|$)", expand=False)
                .astype("string").str.upper()
            )
            type_from_name = type_from_name.combine_first(sf_type)
        # Also check "deliverable" column present in some older bare-PDD files
        if "deliverable" in df.columns:
            deliv_type = (
                df["deliverable"].astype("string")
                .str.extract(r"(?i)(?:^|_)(FSD|CRD|PDD)(?:_|$)", expand=False)
                .astype("string").str.upper()
            )
            type_from_name = type_from_name.combine_first(deliv_type)
        inferred = (type_from_name + "-" + bare_c).where(
            type_from_name.notna() & bare_c.notna()
        )
        crit_id = crit_id.combine_first(inferred)

    # --- Review Date (kept as raw string) ---
    def _format_review_raw(v) -> object:
        if pd.isna(v):
            return pd.NA
        if isinstance(v, pd.Timestamp):
            return v.strftime("%Y-%m-%d")
        s = str(v).strip()
        if s.lower() in ("nat", "nan", "none", ""):
            return pd.NA
        return s

    review_out = pd.Series([pd.NA] * len(df), dtype="object", index=df.index)
    if review_ts_src:
        review_out = df[review_ts_src].apply(_format_review_raw)
    if timestamp_src:
        review_out = review_out.combine_first(df[timestamp_src].apply(_format_review_raw))
    if review_date_src:
        review_out = review_out.combine_first(df[review_date_src].apply(_format_review_raw))
    # Last resort: extract date from source filename (e.g. "...2026-03-03_V1.xlsx")
    if review_out.isna().any() and source_file_src:
        sf = df[source_file_src].astype("string")
        fname_date = sf.str.extract(r"(\d{4}-\d{2}-\d{2})", expand=False)
        fname_date_iso = fname_date.where(
            fname_date.isna(),
            fname_date + "T00:00:00.000Z",
        )
        review_out = review_out.combine_first(fname_date_iso)

    # --- Filename-based Name/DVN fallback (for CRD files that have no Name/DVN column) ---
    # Build ISO date string from review_out (e.g. "2026-03-10" → "2026-03-10T00:00:00.000Z")
    def _review_to_iso(v) -> object:
        if pd.isna(v) or str(v).strip().lower() in ("", "nat", "nan", "none"):
            return pd.NA
        s = (
            str(v).strip()
            .replace("‑", "-")  # non-breaking hyphen
            .replace("–", "-")  # en dash
            .replace("—", "-")  # em dash
        )
        # Already full ISO
        if "T" in s:
            return s
        # YYYY-MM-DD → append time
        try:
            dt = pd.to_datetime(s, errors="coerce")
            if pd.notna(dt):
                return dt.strftime("%Y-%m-%dT00:00:00.000Z")
        except Exception:
            pass
        return pd.NA

    review_iso_full = review_out.apply(_review_to_iso).astype("string")

    # Construct filename-based Name and DVN for rows without a source Name/DVN column.
    # Uses the same prefix-style rules as to_qfs_final_schema:
    #   DVN  = sf_clean + "_V1.0_" + date_iso  (if no version in filename)
    #        = sf_clean + "_" + date_iso         (if version already present)
    #   Name = sf_clean + "_" + date_iso + "_" + crit_id
    if source_file_src:
        sf_fname = (
            df[source_file_src].astype("string")
            .str.replace("%20", " ", regex=False)
            .str.replace(r"\.xlsx?$", "", regex=True)           # strip outer extension
            .str.replace(r"_\d{1,2}-[A-Za-z]{3}-\d{4}$", "", regex=True)  # human date
            .str.replace(r"_[A-Za-z]{3}-\d{1,2}-\d{4}$", "", regex=True)
            .str.replace(r"_\d{4}-\d{2}-\d{2}(\.xlsx)?$", r"\1", regex=True)  # YYYY-MM-DD
        )
        fname_has_v = sf_fname.str.contains(r"_[Vv]\d+(?:\.\d+)?", regex=True, na=False)
        fname_dvn_base = sf_fname.str.replace(r"_v(\d+(?:\.\d+)?)\b", r"_V\1", regex=True)
        fname_dvn = fname_dvn_base.where(
            fname_has_v,
            fname_dvn_base + "_V1.0",
        ) + "_" + review_iso_full
        fname_name_base = sf_fname + "_" + review_iso_full
    else:
        fname_dvn = pd.Series([pd.NA] * len(df), dtype="string", index=df.index)
        fname_name_base = pd.Series([pd.NA] * len(df), dtype="string", index=df.index)

    # --- Name: use source Name (cleaned), fall back to filename-derived name ---
    name_series = (
        df[name_src].astype("string") if name_src
        else pd.Series([pd.NA] * len(df), dtype="string", index=df.index)
    )
    name_out = name_series.apply(remove_human_date).astype("string")

    # For rows where source Name is NA, use filename-based name base
    name_out = name_out.combine_first(fname_name_base.astype("string"))

    # Append criteria ID if Name does not already end with it (e.g. _C001 or _FSD_C001 etc.)
    _crit_end_pat = r"(?:[_\-\u2011\u2013])(?:[A-Z]{2,}[\-_\u2011\u2013])?C\d+\s*$"
    ends_with_crit = name_out.str.contains(_crit_end_pat, regex=True, na=False)
    needs_suffix = ~ends_with_crit & crit_id.notna()
    name_out = name_out.where(~needs_suffix, name_out + "_" + crit_id.astype("string"))

    # --- Deliverable VersionName: use source DVN (cleaned), fall back to filename-derived DVN ---
    dvn_series = (
        df[dvn_src].astype("string") if dvn_src
        else pd.Series([pd.NA] * len(df), dtype="string", index=df.index)
    )
    dvn_out = dvn_series.apply(remove_human_date).astype("string")

    # For rows where source DVN is NA, use filename-based DVN
    dvn_out = dvn_out.combine_first(fname_dvn.astype("string"))

    # Append ISO timestamp to DVN when it doesn't already contain one (existing rows)
    has_iso_ts = dvn_out.str.contains(r"\d{4}-\d{2}-\d{2}T", regex=True, na=False)
    needs_ts = ~has_iso_ts & review_out.notna()
    dvn_out = dvn_out.where(~needs_ts, dvn_out + "_" + review_iso_full.astype("string"))

    # --- Detailed comments ---
    detailed = df[detailed_src] if detailed_src else pd.Series([pd.NA] * len(df), index=df.index)
    if isinstance(detailed, pd.Series) and "comments" in df.columns:
        detailed = detailed.combine_first(df["comments"])

    # --- Build output with positional assignment to avoid index misalignment ---
    out = pd.DataFrame(index=df.index)
    out["Name"] = name_out.to_numpy()
    out["Criteria ID Number – Deliverable Name"] = crit_id.to_numpy()
    out["Criteria Name"] = (
        df[criteria_name_src] if criteria_name_src
        else pd.Series([pd.NA] * len(df), index=df.index)
    ).to_numpy()
    out["Review Date and Timestamp"] = review_out.to_numpy()
    out["Status"] = (
        df[status_src] if status_src else pd.Series([pd.NA] * len(df), index=df.index)
    ).to_numpy()
    out["Score"] = (
        df[score_src] if score_src else pd.Series([pd.NA] * len(df), index=df.index)
    ).to_numpy()
    out["Deliverable VersionName"] = dvn_out.to_numpy()
    out["Detailed comments"] = (
        detailed.to_numpy() if isinstance(detailed, pd.Series)
        else pd.Series([pd.NA] * len(df), index=df.index).to_numpy()
    )

    out = out[P2F_FINAL_COLUMNS]

    # Drop non-data header/footer rows.
    # Case 1: all identifying fields AND score are blank (pure empty rows).
    mask_header = (
        out["Score"].isna()
        & out["Criteria ID Number – Deliverable Name"].isna()
        & out["Criteria Name"].isna()
        & out["Status"].isna()
        & out["Detailed comments"].isna()
    )
    # Case 2: summary/total rows — Score is a aggregated fraction string (e.g. "20.5/31")
    # or review date is clearly non-date text (e.g. "Before Revision").
    # These are footer rows from source spreadsheets, identifiable by having no criteria identity.
    score_num = pd.to_numeric(out["Score"], errors="coerce")
    mask_summary_numeric = (
        out["Criteria ID Number – Deliverable Name"].isna()
        & out["Criteria Name"].isna()
        & out["Name"].isna()
        & out["Deliverable VersionName"].isna()
        & score_num.notna()
        & ~score_num.isin([0.0, 0.5, 1.0])
    )
    mask_summary = (
        out["Criteria ID Number – Deliverable Name"].isna()
        & out["Criteria Name"].isna()
        & (
            out["Score"].astype("string").str.contains(r"\d+/\d+", regex=True, na=False)
            | out["Review Date and Timestamp"].astype("string").str.contains(
                r"(?i)before|after|revision|total|summary", regex=True, na=False
            )
        )
    )
    drop_mask = mask_header | mask_summary | mask_summary_numeric
    if drop_mask.any():
        out = out.loc[~drop_mask].reset_index(drop=True)

    # Strip time portion from Review Date and Timestamp — keep date only (YYYY-MM-DD).
    # This runs last so intermediate processing can still use the full timestamp.
    out["Review Date and Timestamp"] = (
        out["Review Date and Timestamp"]
        .astype("string")
        .str.extract(r"(\d{4}-\d{2}-\d{2})", expand=False)
        .where(out["Review Date and Timestamp"].astype("string").str.contains(r"\d{4}-\d{2}-\d{2}", regex=True, na=False),
               out["Review Date and Timestamp"].astype("string"))
    )

    return out


def _decode_filename(s: pd.Series) -> pd.Series:
    # Handle URL-encoded spaces from some exports.
    return s.astype("string").str.replace("%20", " ", regex=False)


def to_qfs_final_schema(df: pd.DataFrame) -> pd.DataFrame:
    """
    Reshape Q&FS consolidated output into the exact 9-column schema
    used by `output/Q&FS_cleaned_final.xlsx`.
    """
    df = df.copy()
    df = fix_out_of_place_columns(df)

    # Drop obvious blank/footer rows (these are in combined_output-11).
    score_src = _first_present(df, ["score", "Score"])
    if score_src:
        df = df.loc[df[score_src].notna()].reset_index(drop=True)
        # Keep only valid scores (reference uses only 0, 0.5, 1).
        score_num = pd.to_numeric(df[score_src], errors="coerce")
        df = df.loc[score_num.isin([0.0, 0.5, 1.0])].reset_index(drop=True)
    else:
        df = df.reset_index(drop=True)

    # Source columns (snake_case output)
    review_src = _first_present(df, ["review_date", "Review Date"])
    status_src = _first_present(df, ["status", "Status"])
    score_src = _first_present(df, ["score", "Score"])
    crit_name_src = _first_present(df, ["criteria_name", "Criteria Name"])
    detailed_src = _first_present(df, ["detailed_comments", "Detailed Comments", "Detailed comments"])
    comment_src = _first_present(df, ["comment", "Comment"])

    # Criteria ID might be in multiple places; compute it via a single best-effort pass
    # to avoid corner-case misses.
    def _extract_id_from_series(s: pd.Series) -> pd.Series:
        s = (
            s.astype("string")
            .str.replace("–", "-", regex=False)
            .str.replace("—", "-", regex=False)
            .str.replace("‑", "-", regex=False)
        )
        # Accept canonical TYPE-Cnnn first, then the bare-number variant TYPE-nn
        # used by some FSD files (e.g. "FSD-01 - Daily Results Report").
        canonical = s.str.extract(r"(?i)\b([A-Z]{2,}-C\d{1,3})\b", expand=False)
        bare = s.str.extract(r"(?i)\b([A-Z]{2,}-\d{1,3})\b", expand=False)
        # Normalize both branches to the canonical TYPE-Cnnn form
        merged = canonical.combine_first(bare)
        return merged.apply(_normalize_crit_id_token).astype("string")

    # Deliverable VersionName and Name are derived from source filename + review date.
    source_file = _first_present(df, ["_source_file", "source_file"])
    sf_raw = (
        _decode_filename(df[source_file]) if source_file
        else pd.Series([pd.NA] * len(df), dtype="string", index=df.index)
    )

    # Review date → ISO string "YYYY-MM-DDT00:00:00.000Z"
    def _date_to_iso(v) -> object:
        if pd.isna(v) or str(v).strip().lower() in ("", "nat", "nan", "none"):
            return pd.NA
        s = (
            str(v).strip()
            .replace("‑", "-")  # non-breaking hyphen
            .replace("–", "-")  # en dash
            .replace("—", "-")  # em dash
        )
        try:
            for dayfirst in [False, True]:
                dt = pd.to_datetime(s, dayfirst=dayfirst, errors="coerce")
                if pd.notna(dt):
                    return dt.strftime("%Y-%m-%dT00:00:00.000Z")
        except Exception:
            pass
        return pd.NA

    review_dt_src = df[review_src] if review_src else pd.Series([pd.NA] * len(df), index=df.index)
    date_iso = review_dt_src.apply(_date_to_iso).astype("string")

    # Strip outer .xlsx extension only (double .xlsx.xlsx → keeps inner .xlsx)
    sf_base = sf_raw.astype("string").str.replace(r"\.xlsx?$", "", regex=True)

    # Strip trailing human-readable dates; compact dates like 20260309 are kept.
    # The (\.xlsx)? captures an inner .xlsx that may follow the date (double-ext files).
    sf_clean = (
        sf_base
        .str.replace(r"_\d{1,2}-[A-Za-z]{3}-\d{4}$", "", regex=True)       # _07-Mar-2026, _13-FEB-2026
        .str.replace(r"_[A-Za-z]{3}-\d{1,2}-\d{4}$", "", regex=True)       # _Feb-25-2026
        .str.replace(r"_\d{4}-\d{2}-\d{2}(\.xlsx)?$", r"\1", regex=True)   # _2026-03-10 or _2026-03-10.xlsx
    )

    # Name base = sf_clean + "_" + date_iso  (preserves original filename order & version case)
    name_base = sf_clean + "_" + date_iso

    # DVN construction — normalize version case then apply ordering rules
    sf_dvn = sf_clean.str.replace(r"_v(\d+(?:\.\d+)?)\b", r"_V\1", regex=True)

    # Case A – Token style: version immediately before _QualityReview (e.g. _V1.0_QualityReview)
    is_token = sf_dvn.str.contains(r"_V\d+(?:\.\d+)?_QualityReview", regex=True, na=False)
    token_dvn = (
        sf_dvn.str.replace(r"(_V\d+(?:\.\d+)?)(_QualityReview)", r"\2\1", regex=True)
        + "_" + date_iso
    )

    # Case B – Compact date after version (e.g. _V1.0_20260309 at end)
    is_compact = (
        sf_dvn.str.contains(r"_V\d+(?:\.\d+)?_\d{8}$", regex=True, na=False) & ~is_token
    )
    compact_dvn = (
        sf_dvn.str.replace(r"(_V\d+(?:\.\d+)?)(_\d{8})$", r"\2\1", regex=True)
        + "_" + date_iso
    )

    # Case C – Double-extension: ends with _V\d+.xlsx (inner .xlsx sits after version)
    is_double_ext = (
        sf_dvn.str.contains(r"_V\d+(?:\.\d+)?\.xlsx$", regex=True, na=False)
        & ~is_token & ~is_compact
    )
    double_ext_dvn = (
        sf_dvn.str.replace(r"(_V\d+(?:\.\d+)?)(\.xlsx)$", r"\2\1", regex=True)
        + "_" + date_iso
    )

    # Case D – Regular with version (no reordering needed)
    has_v = sf_dvn.str.contains(r"_V\d+(?:\.\d+)?", regex=True, na=False)
    is_regular_v = has_v & ~is_token & ~is_compact & ~is_double_ext
    regular_v_dvn = sf_dvn + "_" + date_iso

    # Case E – No version present: insert _V1.0 before date
    no_v_dvn = sf_dvn + "_V1.0_" + date_iso

    dvn = token_dvn.where(
        is_token,
        compact_dvn.where(
            is_compact,
            double_ext_dvn.where(
                is_double_ext,
                regular_v_dvn.where(is_regular_v, no_v_dvn),
            ),
        ),
    )

    # Criteria ID extraction
    crit_id = pd.Series([pd.NA] * len(df), index=df.index, dtype="string")
    for col in ["criteria_id", "Criteria ID", "criteria_id_deliverable_name", "nt", "NT", "deliverable_name", "Deliverable Name", "1"]:
        if col in df.columns:
            crit_id = crit_id.combine_first(_extract_id_from_series(df[col]))
    crit_id = crit_id.astype("string").str.upper().reset_index(drop=True)

    # Build output with a stable index and assign positionally to avoid alignment bugs.
    out = pd.DataFrame(index=df.index)
    review_out = pd.to_datetime(review_dt_src, errors="coerce")
    out["Review Date"] = review_out.dt.date.to_numpy()
    out["Status"] = (df[status_src] if status_src else pd.Series([pd.NA] * len(df))).to_numpy()
    out["Score"] = (df[score_src] if score_src else pd.Series([pd.NA] * len(df))).to_numpy()
    out["Criteria Name"] = (df[crit_name_src] if crit_name_src else pd.Series([pd.NA] * len(df))).to_numpy()
    out["Criteria ID"] = crit_id.to_numpy()

    detailed = df[detailed_src] if detailed_src else pd.NA
    if isinstance(detailed, pd.Series) and "comments" in df.columns:
        detailed = detailed.combine_first(df["comments"])
    if isinstance(detailed, pd.Series):
        out["Detailed Comments"] = detailed.to_numpy()
    else:
        out["Detailed Comments"] = pd.Series([pd.NA] * len(df)).to_numpy()

    comment = df[comment_src] if comment_src else pd.NA
    if isinstance(comment, pd.Series) and "comments" in df.columns:
        comment = comment.combine_first(df["comments"])
    if isinstance(comment, pd.Series):
        out["Comment"] = comment.to_numpy()
    else:
        out["Comment"] = pd.Series([pd.NA] * len(df)).to_numpy()

    out["Deliverable VersionName"] = dvn.reset_index(drop=True).to_numpy()

    # Name suffix: use Criteria ID when extracted from a proper criteria column;
    # fall back to Criteria Name for non-standard source columns (nt / 1).
    std_crit_mask = pd.Series([False] * len(df), index=df.index)
    for _col in ["criteria_id", "Criteria ID", "criteria_id_deliverable_name"]:
        if _col in df.columns:
            std_crit_mask = std_crit_mask | _extract_id_from_series(df[_col]).notna()
    name_suffix = crit_id.astype("string")
    if crit_name_src:
        _crit_name_vals = df[crit_name_src].astype("string").reset_index(drop=True)
        name_suffix = _crit_name_vals.where(~std_crit_mask.reset_index(drop=True), name_suffix.reset_index(drop=True))
    out["Name"] = (name_base + "_" + name_suffix).reset_index(drop=True).to_numpy()

    # Reorder and ensure exact columns
    out = out[QFS_FINAL_COLUMNS]

    # Drop any remaining rows that still don't have a score (safety)
    out = out.loc[out["Score"].notna()].reset_index(drop=True)
    return out

def clean_combined_file(input_file: str, output_file: str) -> dict:
    df = pd.read_excel(input_file)
    print(f"Loaded {len(df)} rows from {input_file}")
    print(f"Columns: {list(df.columns)}")

    # Fix common consolidation issues for snake_case consolidated outputs
    df = fix_out_of_place_columns(df)

    if "Review Date" in df.columns:
        df["Review Date"] = df["Review Date"].apply(standardize_date)
        print("✓ Review Date standardized to YYYY-MM-DD")
    else:
        print("WARNING: 'Review Date' column not found")

    if "Name" in df.columns:
        df["Name"] = df["Name"].apply(remove_human_date)
        print("✓ Name: human-readable date removed, ISO timestamp kept")
    else:
        print("WARNING: 'Name' column not found")

    if "Deliverable VersionName" in df.columns:
        df["Deliverable VersionName"] = df["Deliverable VersionName"].apply(remove_human_date)
        print("✓ Deliverable VersionName: human-readable date removed, ISO timestamp kept")
    else:
        print("WARNING: 'Deliverable VersionName' column not found")

    if "Criteria ID" in df.columns:
        df["Criteria ID"] = df["Criteria ID"].apply(clean_criteria_id)
        print("✓ Criteria ID cleaned (extracted ID only)")
    else:
        print("WARNING: 'Criteria ID' column not found")

    # If the dataset looks like P2F or Q&FS, reshape to the reference schema.
    if any(c in df.columns for c in ["criteria_id_number_deliverable_name", "Criteria ID Number – Deliverable Name"]):
        df_out = to_p2f_final_schema(df)
    elif any(c in df.columns for c in ["criteria_id_deliverable_name", "Detailed Comments", "detailed_comments"]) and any(
        c in df.columns for c in ["Review Date", "review_date"]
    ):
        # Likely Q&FS style export
        df_out = to_qfs_final_schema(df)
    else:
        df_out = df

    df_out.to_excel(output_file, index=False)
    print(f"\n✓ Saved cleaned file to: {output_file}")
    return {"input_file": input_file, "output_file": output_file, "rows": len(df_out), "cols": len(df_out.columns)}


def main():
    parser = argparse.ArgumentParser(description="Clean consolidated output (dates, IDs, name fields).")
    parser.add_argument("--input", default=DEFAULT_INPUT_FILE, help="Input combined .xlsx file")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_FILE, help="Output cleaned .xlsx file")
    args = parser.parse_args()
    clean_combined_file(args.input, args.output)


if __name__ == "__main__":
    main()