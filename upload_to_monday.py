#!/usr/bin/env python3
"""
Monday.com upload automation for the Excel Pipeline.

USAGE
-----
Step 1 – Discover board structure (groups + column IDs):
    python upload_to_monday.py --discover
    python upload_to_monday.py --discover --board 225270003

Step 2 – Fill in monday_config.json with the IDs printed by --discover.

Step 3 – Upload:
    python upload_to_monday.py --upload --folder /path/to/data/P2F
    python upload_to_monday.py --upload --folder /path/to/data --type qfs
    python upload_to_monday.py --upload --file /path/to/output.xlsx

Optional flags:
    --dry-run     Print what would be uploaded without actually sending
    --type        Force dataset type: p2f | qfs | auto (default: auto)
"""

import sys
import os
import re
import json
import time
import argparse
import logging
import glob
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import pandas as pd
from typing import Optional, Tuple
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ──────────────────────────────────────────────────────────────
# Paths
# ──────────────────────────────────────────────────────────────
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH  = os.path.join(SCRIPT_DIR, "monday_config.json")
API_URL      = "https://api.monday.com/v2"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

_SESSION_LOCK = threading.Lock()
_SESSION_BY_TOKEN: dict[str, requests.Session] = {}


# ──────────────────────────────────────────────────────────────
# Config helpers
# ──────────────────────────────────────────────────────────────

def load_config() -> dict:
    """
    Load and validate monday_config.json from the same directory as this script.

    Exits immediately with a clear error message if:
    - The config file does not exist at all.
    - The api_token field is missing or still contains the placeholder value.

    Returns the full config dict on success. All other functions receive this
    dict as their `cfg` parameter rather than reading the file themselves.
    """
    if not os.path.exists(CONFIG_PATH):
        sys.exit(f"Config file not found: {CONFIG_PATH}\n"
                 "Run 'python upload_to_monday.py --help' for setup instructions.")
    with open(CONFIG_PATH) as f:
        cfg = json.load(f)
    token = cfg.get("api_token", "")
    if not token or token == "YOUR_API_TOKEN_HERE":
        sys.exit("ERROR: Please set your Monday.com API token in monday_config.json")
    return cfg


# ──────────────────────────────────────────────────────────────
# Monday.com GraphQL helpers
# ──────────────────────────────────────────────────────────────

def _gql(token: str, query: str, variables: dict = None) -> dict:
    """
    Execute a single GraphQL query or mutation against the Monday.com API.

    All network calls in this file go through this one function so that
    headers (auth, API version) and error handling are applied consistently.

    Raises:
        requests.HTTPError  – if the HTTP response status is 4xx / 5xx.
        RuntimeError        – if Monday.com returns a JSON-level 'errors' field
                              (e.g. invalid field name, permission denied).

    Returns the 'data' dict from the response, or {} if the key is absent.
    Variables are forwarded as a separate JSON payload key so that Monday.com
    can cache the query shape independently of its runtime values.
    """
    headers = {
        "Authorization": token,
        "Content-Type": "application/json",
        "API-Version": "2024-01",
    }
    payload = {"query": query}
    if variables:
        payload["variables"] = variables
    resp = _get_session(token).post(
        API_URL,
        headers=headers,
        json=payload,
        timeout=(5, 60),
    )
    resp.raise_for_status()
    data = resp.json()
    if "errors" in data:
        raise RuntimeError(f"Monday.com API error: {data['errors']}")
    return data.get("data", {})


def _get_session(token: str) -> requests.Session:
    """
    Return a cached Session per token with connection pooling + retries.
    """
    with _SESSION_LOCK:
        session = _SESSION_BY_TOKEN.get(token)
        if session is not None:
            return session

        session = requests.Session()
        retry = Retry(
            total=5,
            connect=5,
            read=5,
            backoff_factor=0.5,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=frozenset(["POST"]),
            raise_on_status=False,
            respect_retry_after_header=True,
        )
        adapter = HTTPAdapter(pool_connections=50, pool_maxsize=50, max_retries=retry)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        _SESSION_BY_TOKEN[token] = session
        return session


def fetch_board_info(token: str, board_id: str) -> dict:
    """
    Fetch a board's full structure: its name, all group IDs/titles, and all
    column IDs/titles/types.

    Used by the `discover` command to print a human-readable summary so the
    user can copy the right IDs into monday_config.json.
    Also used by upload_criteria_names_only() to find the first available group
    on the criteria-name board without needing it pre-configured.

    Exits if the board is not accessible (wrong ID or insufficient permissions).
    Returns the first (and only) board dict from the API response.
    """
    query = """
    query ($board_id: [ID!]!) {
      boards(ids: $board_id) {
        id name
        groups { id title }
        columns { id title type }
      }
    }
    """
    data = _gql(token, query, {"board_id": [board_id]})
    boards = data.get("boards", [])
    if not boards:
        sys.exit(f"Board {board_id} not found. Check your board ID and API token permissions.")
    return boards[0]


def create_item(token: str, board_id: str, group_id: str, item_name: str,
                column_values: dict, dry_run: bool = False) -> Optional[str]:
    """
    Create a single item on a Monday.com board and return its new item ID.

    `column_values` must be a dict of {monday_col_id: formatted_value} where
    each value is already shaped for the Monday.com API (e.g. {"date": "2026-03-19"}
    for date columns, {"label": "Done"} for status columns, a plain string for
    text columns, etc.). The dict is JSON-serialised before being sent.

    In dry-run mode logs what would be created and returns None immediately
    without making any API call.

    Returns the new item's string ID on success, or None on dry-run.
    """
    if dry_run:
        log.info(f"[DRY-RUN] Would create item: {item_name!r} in group {group_id}")
        return None
    query = """
    mutation ($board_id: ID!, $group_id: String!, $item_name: String!, $col_vals: JSON!) {
      create_item(
        board_id: $board_id,
        group_id: $group_id,
        item_name: $item_name,
        column_values: $col_vals
      ) { id }
    }
    """
    data = _gql(token, query, {
        "board_id":  board_id,
        "group_id":  group_id,
        "item_name": item_name,
        "col_vals":  json.dumps(column_values),
    })
    return data.get("create_item", {}).get("id")


def fetch_existing_item_names(token: str, board_id: str, group_id: str) -> set:
    """
    Return a set of all existing item names (titles) in a specific board group.

    Used before uploading to the criteria-name board so that already-present
    criteria names are skipped rather than duplicated.
    Returns an empty set on any API error so that the upload can still proceed
    (worst case: a duplicate is created rather than the whole upload failing).

    Note: the items_page query is limited to the first page (~100 items by default
    in Monday's API). For very large groups this may miss items on subsequent pages,
    but for typical criteria-name boards this is sufficient.
    """
    query = """
    query ($board_id: [ID!]!, $group_id: [String!]!) {
      boards(ids: $board_id) {
        groups(ids: $group_id) {
          items_page { items { name } }
        }
      }
    }
    """
    try:
        data = _gql(token, query, {"board_id": [board_id], "group_id": [group_id]})
        groups = data.get("boards", [{}])[0].get("groups", [])
        if groups:
            return {i["name"] for i in groups[0].get("items_page", {}).get("items", [])}
    except Exception:
        pass
    return set()


# ──────────────────────────────────────────────────────────────
# Pipeline helpers (reuse existing pipeline)
# ──────────────────────────────────────────────────────────────

def _nc(c: str) -> str:
    """
    Normalise a column name to a safe, lowercase, underscore-separated identifier.

    Strips leading/trailing whitespace, lowercases, then replaces every run of
    non-word characters (spaces, hyphens, dashes, en-dashes, etc.) with a single
    underscore. Falls back to "col" if the result would be empty.

    Used by run_pipeline() to make raw Excel column headers safe for pandas
    attribute access and consistent string matching across different source files.
    """
    s = str(c).strip().lower()
    s = re.sub(r"[^\w]+", "_", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "col"


def _dedup_cols(cols):
    """
    Ensure every column name in a list is unique by appending _2, _3, … to duplicates.

    pandas raises an error when a DataFrame has duplicate column names, which can
    happen when multiple source Excel files have identically named columns that get
    combined. This function guarantees uniqueness while preserving the original order
    and the first occurrence's name unchanged.
    """
    used = set(); seen = {}; out = []
    for c in cols:
        base = str(c)
        if base not in used:
            used.add(base); seen[base] = 1; out.append(base)
        else:
            n = seen.get(base, 0) + 1
            cand = f"{base}_{n}"
            while cand in used:
                n += 1; cand = f"{base}_{n}"
            seen[base] = n; used.add(cand); out.append(cand)
    return out


def run_pipeline(folder: str, dataset_type: str = "auto") -> Tuple[pd.DataFrame, str]:
    """
    Run the full Excel processing pipeline on a folder of source files and return
    a clean, schema-normalised DataFrame ready for upload.

    Steps:
    1. combine_folder_to_frames() merges all .xlsx files in `folder` into one DataFrame.
    2. Column names are normalised (_nc) and deduplicated (_dedup_cols) so that
       downstream string matching is reliable regardless of source file formatting.
    3. Fully empty rows are dropped.
    4. Dataset type is auto-detected from column names and source file names if
       `dataset_type` is "auto":
       - P2F  → has a "Criteria ID Number – Deliverable Name" style column, or source
                 files whose names contain "CRD_".
       - Q&FS → has a "review_date" column plus a Criteria ID / Criteria Name column.
       - Falls back to raw combined data with a warning if neither pattern matches.
    5. The appropriate schema transform (to_p2f_final_schema or to_qfs_final_schema)
       is applied to standardise column names and add any derived columns.

    Returns (df, dataset_type) where dataset_type is the resolved string
    ("p2f", "qfs", or "auto" if detection failed).
    """
    sys.path.insert(0, SCRIPT_DIR)
    from src.combiner import combine_folder_to_frames
    from src.data_q2  import to_p2f_final_schema, to_qfs_final_schema

    log.info(f"Running pipeline on: {folder}")
    result   = combine_folder_to_frames(folder)
    combined = result["combined_df"]
    combined.columns = _dedup_cols([_nc(c) for c in combined.columns])
    combined = combined.dropna(how="all")

    if dataset_type == "auto":
        has_num_col = any(c in combined.columns for c in [
            "criteria_id_number_deliverable_name",
            "Criteria ID Number \u2013 Deliverable Name",
        ])
        has_crd_col = "criteria_id_deliverable_name" in combined.columns
        src_col     = next((c for c in ["source_file", "_source_file"] if c in combined.columns), None)
        src_is_crd  = (
            has_crd_col and src_col is not None
            and combined[src_col].astype(str).str.contains(r"(?i)CRD_", regex=True, na=False).any()
        )
        is_p2f = has_num_col or src_is_crd

        if is_p2f:
            dataset_type = "p2f"
        elif (
            "review_date" in combined.columns
            and any(c in combined.columns for c in [
                "criteria_id_deliverable_name", "criteria_id", "criteria_name", "nt", "1"
            ])
        ):
            dataset_type = "qfs"

    if dataset_type == "p2f":
        log.info("Detected P2F dataset — applying P2F schema")
        combined = to_p2f_final_schema(combined)
    elif dataset_type == "qfs":
        log.info("Detected Q&FS dataset — applying Q&FS schema")
        combined = to_qfs_final_schema(combined)
    else:
        log.warning("Could not detect dataset type. Uploading raw combined data.")

    log.info(f"Pipeline produced {len(combined)} rows, {len(combined.columns)} columns")
    return combined, dataset_type


def load_from_file(filepath: str) -> Tuple[pd.DataFrame, str]:
    """
    Load a previously-processed output .xlsx file instead of re-running the pipeline.

    Prefers the sheet named "Full Data " (note trailing space, as produced by the
    pipeline's export) then "Full Data", then falls back to the first sheet in the
    workbook if neither named sheet exists.

    Dataset type is auto-detected from column headers:
    - P2F  → any column whose name contains "Criteria ID Number".
    - Q&FS → everything else.

    This allows the --file CLI flag to be used when the pipeline has already been
    run separately and only the Monday upload step needs to be re-run.

    Returns (df, dataset_type).
    """
    log.info(f"Loading from file: {filepath}")
    xl = pd.ExcelFile(filepath)
    sheet = "Full Data " if "Full Data " in xl.sheet_names else (
            "Full Data"  if "Full Data"  in xl.sheet_names else xl.sheet_names[0])
    df = pd.read_excel(filepath, sheet_name=sheet)
    log.info(f"Loaded {len(df)} rows from sheet '{sheet}'")

    # Auto-detect type from columns
    dataset_type = "p2f" if any("Criteria ID Number" in str(c) for c in df.columns) else "qfs"
    return df, dataset_type


# ──────────────────────────────────────────────────────────────
# Column value builder for Monday.com API
# ──────────────────────────────────────────────────────────────

def _build_col_values(row: pd.Series, col_map: dict, col_types: dict) -> dict:
    """
    Build the column_values JSON for a Monday.com item.
    col_types: {monday_col_id: monday_col_type}  (from board discovery)
    """
    values = {}
    for excel_col, monday_col_id in col_map.items():
        if monday_col_id in ("name", "_comment"):
            continue
        if excel_col not in row.index:
            continue
        raw = row[excel_col]
        if pd.isna(raw):
            continue

        val_str  = str(raw).strip()
        col_type = col_types.get(monday_col_id, "text")

        if col_type == "text" or col_type == "long_text":
            values[monday_col_id] = {"text": val_str}
        elif col_type == "numbers":
            try:
                values[monday_col_id] = str(float(val_str))
            except ValueError:
                values[monday_col_id] = {"text": val_str}
        elif col_type == "date":
            # Extract YYYY-MM-DD from any date string
            m = re.search(r"(\d{4}-\d{2}-\d{2})", val_str)
            if m:
                values[monday_col_id] = {"date": m.group(1)}
        elif col_type == "status":
            values[monday_col_id] = {"label": val_str}
        elif col_type == "dropdown":
            values[monday_col_id] = {"labels": [val_str]}
        else:
            values[monday_col_id] = val_str

    return values


# ──────────────────────────────────────────────────────────────
# Determine which Monday.com group a row belongs to
# ──────────────────────────────────────────────────────────────

def _resolve_group(row: pd.Series, dataset_type: str, group_map: dict) -> str:
    """
    Route a row to the correct Monday.com group based on its Criteria ID prefix.

    group_map keys must be  "FSD", "CRD", "PDD"  (matching the first 3 letters of
    the Criteria ID Number column, e.g. FSD_C001 → FSD group).
    Falls back to the first available group if no prefix match is found.
    """
    # Find the Criteria ID column in this row
    crit_col = next(
        (c for c in row.index
         if "Criteria ID" in str(c) and "Number" in str(c)),
        next((c for c in row.index if "Criteria ID" in str(c)), None)
    )
    crit_val = str(row[crit_col]).strip() if crit_col and pd.notna(row.get(crit_col)) else ""

    # Match prefix FSD / CRD / PDD directly to the group_map key
    for prefix in ("FSD", "CRD", "PDD"):
        if crit_val.upper().startswith(prefix) and prefix in group_map:
            gid = group_map[prefix]
            if gid.startswith("GROUP_ID"):
                sys.exit(
                    f"ERROR: group_map['{prefix}'] is still a placeholder. "
                    "Run --discover and fill in the real group ID in monday_config.json"
                )
            return gid

    # Fallback: first non-comment key in the map
    first = next((v for k, v in group_map.items() if not k.startswith("_")), None)
    if first and not first.startswith("GROUP_ID"):
        return first

    sys.exit(
        "ERROR: Could not match row to a Monday.com group.\n"
        "Run: python upload_to_monday.py --discover\n"
        "Then fill in group_map in monday_config.json with FSD / CRD / PDD group IDs."
    )


def _run_parallel_batches(tasks: list, batch_size: int, delay: float) -> list:
    """
    Execute callables concurrently in batches, preserving throttling between batches.
    """
    if not tasks:
        return []
    safe_batch = max(1, int(batch_size or 1))
    safe_delay = max(0.0, float(delay or 0.0))
    results = []
    for start in range(0, len(tasks), safe_batch):
        batch = tasks[start:start + safe_batch]
        with ThreadPoolExecutor(max_workers=min(safe_batch, len(batch))) as pool:
            futures = [pool.submit(task) for task in batch]
            for fut in as_completed(futures):
                results.append(fut.result())
        if start + safe_batch < len(tasks) and safe_delay:
            time.sleep(safe_delay)
    return results


# ──────────────────────────────────────────────────────────────
# Upload logic
# ──────────────────────────────────────────────────────────────

def upload_unique_names(df: pd.DataFrame, dataset_type: str, cfg: dict,
                        dry_run: bool = False) -> dict:
    """
    Upload one Monday.com item per unique Deliverable VersionName to the L2 board.

    The L2 board only stores one item per checklist version (DVN), not one per
    criteria row, so this function deduplicates before uploading.

    How it works:
    1. Locate the best item-name column in the DataFrame (DVN preferred, then
       "Name", then a name derived by stripping the _Cxxx suffix from Criteria ID).
    2. Build a deduplicated list of (item_name, first_criteria_id_for_that_item).
    3. For each unique name, resolve the target Monday group using the Criteria ID
       prefix (FSD / CRD / PDD → group_map in config).
    4. Call create_item() with an empty column_values dict — the L2 board only
       allows the item name to be written; all other columns are mirrors/formulas.
    5. Pause between batches (batch_size / delay from config) to avoid rate limits.

    Does NOT skip duplicates — if a DVN already exists on the board a second item
    is created (on_duplicate = "always_create" is the intended behaviour here).

    Returns {"created": int, "skipped": int, "errors": int}.
    """
    token      = cfg["api_token"]
    board_id   = cfg["board_id"]
    group_map  = {k: v for k, v in cfg["group_map"].items() if not k.startswith("_")}
    batch_size = cfg.get("upload", {}).get("batch_size", 10)
    delay      = cfg.get("upload", {}).get("delay_between_batches_seconds", 1)

    # Find the Criteria ID column (used to route to correct group)
    crit_col = next(
        (c for c in df.columns if "Criteria ID" in str(c) and "Number" in str(c)),
        next((c for c in df.columns if "Criteria ID" in str(c)), None)
    )

    # Determine the best item-name column in priority order:
    #   1. Deliverable VersionName  (preferred — unique per checklist version)
    #   2. Name                     (fallback)
    #   3. Strip _Cxxx suffix from Criteria ID Number column (last resort for old files)
    dvn_col = next(
        (c for c in df.columns if "Deliverable VersionName" in str(c)),
        next((c for c in df.columns if "deliverable" in str(c).lower() and "version" in str(c).lower()), None)
    )
    name_col = next((c for c in df.columns if str(c).strip() == "Name"), None)

    dvn_filled  = dvn_col  and df[dvn_col].notna().any()
    name_filled = name_col and df[name_col].notna().any()

    if dvn_filled:
        item_name_col = dvn_col
        log.info(f"Using '{dvn_col}' as item name source")
    elif name_filled:
        item_name_col = name_col
        log.warning(f"'Deliverable VersionName' is empty — falling back to '{name_col}' as item name")
    elif crit_col and df[crit_col].notna().any():
        # Last resort: derive deliverable name from Criteria ID by stripping the _Cxxx suffix
        item_name_col = None
        log.warning(
            "Both 'Deliverable VersionName' and 'Name' are empty in this file.\n"
            "  This output was likely generated before the latest pipeline fixes.\n"
            "  Re-run the pipeline on your source folder to get a fresh output, then upload again:\n"
            "    python upload_to_monday.py upload --folder /path/to/data\n"
            "  Falling back to deriving names from Criteria ID column..."
        )
    else:
        log.error(
            "No item names found — 'Deliverable VersionName', 'Name', and 'Criteria ID' are all empty.\n"
            "Please re-run the pipeline on your source folder first:\n"
            "  python upload_to_monday.py upload --folder /path/to/your/data"
        )
        return {"created": 0, "skipped": 0, "errors": 0}

    # Build a deduplicated list: (item_name, first_criteria_id_seen)
    seen_dvns = {}  # item_name -> first criteria_id for that item
    for _, row in df.iterrows():
        if item_name_col:
            dvn = str(row.get(item_name_col, "")).strip()
        else:
            # Derive from Criteria ID: strip trailing _Cxxx or ‑Cxxx
            raw_crit = str(row.get(crit_col, "")).strip()
            dvn = re.sub(r"[_\-\u2011\u2013]C\d+$", "", raw_crit).strip()

        if not dvn or dvn in ("nan", "<NA>", ""):
            continue
        if dvn not in seen_dvns:
            crit_val = str(row.get(crit_col, "")).strip() if crit_col else ""
            seen_dvns[dvn] = crit_val

    unique_dvns = list(seen_dvns.items())  # [(item_name, criteria_id), ...]
    log.info(f"Found {len(unique_dvns)} unique item names to upload")

    stats = {"created": 0, "skipped": 0, "errors": 0}

    def _make_task(dvn_name: str, crit_id: str):
        # Route to group based on Criteria ID prefix (FSD / CRD / PDD)
        group_id = None
        matched_prefix = "?"
        for prefix in ("FSD", "CRD", "PDD"):
            if crit_id.upper().startswith(prefix) and prefix in group_map:
                group_id = group_map[prefix]
                matched_prefix = prefix
                break
        if not group_id:
            group_id = next(iter(group_map.values()))
            log.warning(f"  Could not determine group for {dvn_name!r} "
                        f"(crit_id={crit_id!r}) — using first group")

        def _task():
            try:
                item_id = create_item(token, board_id, group_id, dvn_name, {}, dry_run)
                return {"ok": True, "item_id": item_id, "name": dvn_name, "prefix": matched_prefix}
            except Exception as e:
                return {"ok": False, "name": dvn_name, "error": str(e)}
        return _task

    tasks = [_make_task(dvn_name, crit_id) for dvn_name, crit_id in unique_dvns]
    for result in _run_parallel_batches(tasks, batch_size=batch_size, delay=delay):
        if result["ok"]:
            if result.get("item_id"):
                log.info(f"  ✓ [{result['prefix']}] Created item {result['item_id']}: {result['name'][:65]!r}")
            stats["created"] += 1
        else:
            log.error(f"  ✗ Failed: {result['name']!r}: {result['error']}")
            stats["errors"] += 1

    return stats


def upload_criteria_names_only(df: pd.DataFrame, cfg: dict, dry_run: bool = False) -> dict:
    """
    Upload unique Criteria Names to the separate criteria-name-only board.

    This board stores one item per criteria name (not per DVN). It is an optional
    second board configured via `criteria_name_only_board_id` in monday_config.json.
    If that key is absent or still contains the placeholder value, this function
    logs a warning and returns immediately without uploading anything.

    How it works:
    1. Fetch the board's first available group (the target group is not pre-configured;
       any group on that board is acceptable).
    2. Fetch all existing item names in that group to avoid duplicates.
    3. Iterate over unique values in the "Criteria Name" column of the DataFrame,
       skipping any that already exist on the board.
    4. Create one item per new unique criteria name with no extra column values
       (only the item title / name field is populated).

    Returns {"created": int, "skipped": int, "errors": int}.
    """
    token    = cfg["api_token"]
    board_id = cfg.get("criteria_name_only_board_id", "")
    if not board_id or board_id == "SECOND_BOARD_ID_HERE":
        log.warning("criteria_name_only_board_id not set in config — skipping criteria-name board.")
        return {"created": 0, "skipped": 0, "errors": 0}

    col_map    = {k: v for k, v in cfg.get("criteria_name_only_column_map", {}).items()
                  if not k.startswith("_")}
    batch_size = cfg.get("upload", {}).get("batch_size", 10)
    delay      = cfg.get("upload", {}).get("delay_between_batches_seconds", 1)

    # For this board, item name = Criteria Name; fetch first available group
    board  = fetch_board_info(token, board_id)
    groups = board.get("groups", [])
    if not groups:
        log.error("No groups found on criteria-name board.")
        return {"created": 0, "skipped": 0, "errors": 0}
    group_id = groups[0]["id"]

    existing = fetch_existing_item_names(token, board_id, group_id)
    stats    = {"created": 0, "skipped": 0, "errors": 0}

    crit_name_col = next((c for c in df.columns if "Criteria Name" in str(c)), None)
    if not crit_name_col:
        log.error("No 'Criteria Name' column found in output — cannot upload to criteria-name board.")
        return stats

    unique_names = df[crit_name_col].dropna().unique()
    log.info(f"Uploading {len(unique_names)} unique criteria names to board {board_id}...")

    to_create = []
    for raw_name in unique_names:
        name_val = str(raw_name).strip()
        if not name_val or name_val == "nan":
            continue
        if name_val in existing:
            stats["skipped"] += 1
            continue
        to_create.append(name_val)

    def _make_task(name_val: str):
        def _task():
            try:
                item_id = create_item(token, board_id, group_id, name_val, {}, dry_run)
                return {"ok": True, "item_id": item_id, "name": name_val}
            except Exception as e:
                return {"ok": False, "name": name_val, "error": str(e)}
        return _task

    tasks = [_make_task(name_val) for name_val in to_create]
    for result in _run_parallel_batches(tasks, batch_size=batch_size, delay=delay):
        if result["ok"]:
            if result.get("item_id"):
                log.info(f"  ✓ Created: {result['name']!r}")
            stats["created"] += 1
            existing.add(result["name"])
        else:
            log.error(f"  ✗ Failed: {result['name']!r}: {result['error']}")
            stats["errors"] += 1

    return stats


def fetch_criteria_library_id(token: str, l3_board_id: str, relation_col_id: str = "board_relation_mkvv6kmc") -> Optional[str]:
    """
    Auto-discover the Criteria Library Board ID by reading the L3 board's column settings.

    The L3 board has a board_relation column (default ID: "board_relation_mkvv6kmc")
    whose settings_str JSON contains a "boardIds" array pointing to the Criteria
    Library Board. This avoids hardcoding the Criteria Library Board ID in config.

    Returns the board ID as a string if found, or None if the relation column is
    absent, the settings JSON is malformed, or boardIds is empty.
    """
    query = """
    query ($board_id: [ID!]!) {
      boards(ids: $board_id) {
        columns { id settings_str }
      }
    }
    """
    data = _gql(token, query, {"board_id": l3_board_id})
    columns = data.get("boards", [{}])[0].get("columns", [])
    for col in columns:
        if col.get("id") == relation_col_id:
            try:
                settings = json.loads(col.get("settings_str", "{}"))
                board_ids = settings.get("boardIds", [])
                if board_ids:
                    return str(board_ids[0])
            except (json.JSONDecodeError, IndexError):
                pass
    return None


def fetch_criteria_library_map(token: str, criteria_board_id: str) -> dict:
    """
    Build a lookup map of {item_name: item_id} for every item on the Criteria Library Board.

    This map is used during L3 upload to populate the board_relation column on each
    row, linking it back to its corresponding entry in the Criteria Library.

    Uses cursor-based pagination (500 items per page) to handle boards with more
    items than Monday's single-page limit. Continues fetching pages until the API
    returns no further cursor.

    Returns a dict of {stripped_item_name: item_id_string}.
    """
    query = """
    query ($board_id: [ID!]!, $cursor: String) {
      boards(ids: $board_id) {
        items_page(limit: 500, cursor: $cursor) {
          cursor
          items { id name }
        }
      }
    }
    """
    name_to_id: dict = {}
    cursor = None
    while True:
        data = _gql(token, query, {"board_id": criteria_board_id, "cursor": cursor})
        page = data.get("boards", [{}])[0].get("items_page", {})
        for item in page.get("items", []):
            name_to_id[item["name"].strip()] = item["id"]
        cursor = page.get("cursor")
        if not cursor:
            break
    log.info(f"Criteria Library: loaded {len(name_to_id)} items for relation linking")
    return name_to_id


def create_group_on_board(token: str, board_id: str, group_name: str) -> str:
    """
    Create a new group on the given board and return its ID.

    Called once at the start of each L3 upload to create a dated group
    (e.g. "P2F_2026-03-19") that contains all rows from that upload session.
    Creating a fresh group per upload makes it easy to identify when each batch
    was uploaded and to trigger the sync automation on just that batch.

    Returns the new group's ID string, or an empty string if creation failed.
    """
    query = """
    mutation ($board_id: ID!, $group_name: String!) {
      create_group(board_id: $board_id, group_name: $group_name) { id }
    }
    """
    data = _gql(token, query, {"board_id": board_id, "group_name": group_name})
    return data.get("create_group", {}).get("id", "")


def upload_to_l3_board(df: pd.DataFrame, dataset_type: str, cfg: dict, dry_run: bool = False) -> dict:
    """
    Upload every criteria row as an individual item to the L3 board.

    Unlike upload_unique_names() which deduplicates to one item per DVN, this
    function uploads one item per row — the L3 board tracks the detailed scoring
    and review data for each individual criteria.

    How it works:
    1. Resolve column mapping: for each entry in l3_column_map, try the primary
       column name first, then each alias from l3_column_aliases, case-insensitively.
       This makes the upload robust across P2F and Q&FS files that name columns
       slightly differently.
    2. Auto-discover the Criteria Library Board ID from the L3 board's relation
       column settings, then fetch all criteria items for relation-linking.
    3. Create a new dated group on the L3 board (e.g. "P2F_2026-03-19") to
       contain this upload batch.
    4. For each row, build the column_values payload with correct Monday.com types:
       - date columns → {"date": "YYYY-MM-DD"}
       - status columns → {"label": "value"}
       - numeric columns → float
       - long_text columns → {"text": "value"}
       - text columns → plain string
    5. Attempt to link each row to its Criteria Library item via the board_relation
       column using the Criteria ID value (tries both "_" and "-" separators).
    6. Upload in batches with a configurable delay to avoid rate limiting.

    Stores the created group_id in the returned stats dict so the caller can pass
    it to sync_l3_group() to trigger the Monday automation.

    Returns {"created": int, "skipped": int, "errors": int, "group_id": str}.
    """
    token    = cfg["api_token"]
    board_id = cfg.get("l3_board_id", "")
    if not board_id or board_id in ("L3_BOARD_ID_HERE", ""):
        log.warning("l3_board_id not set in config — skipping L3 upload.")
        return {"created": 0, "skipped": 0, "errors": 0}

    from datetime import date as _date

    col_map    = {k: v for k, v in cfg.get("l3_column_map", {}).items()
                  if not k.startswith("_")}
    # Build reverse alias map: monday_col_id → [candidate pipeline col names]
    aliases    = {k: v for k, v in cfg.get("l3_column_aliases", {}).items()
                  if not k.startswith("_")}
    # Pre-resolve which df column to use for each Monday col ID
    # (tries col_map key first, then each alias, all case-insensitive)
    df_cols_lower = {str(c).strip().lower(): c for c in df.columns}
    resolved_cols: dict = {}  # pipeline_col_name → monday_col_id (final effective mapping)
    for pipeline_col, monday_col_id in col_map.items():
        candidates = [pipeline_col] + aliases.get(monday_col_id, [])
        for candidate in candidates:
            matched = df_cols_lower.get(candidate.strip().lower())
            if matched is not None:
                resolved_cols[matched] = monday_col_id
                break
    log.info(f"L3 column mapping resolved: { {k: v for k, v in resolved_cols.items()} }")

    batch_size = cfg.get("upload", {}).get("batch_size", 10)
    delay      = cfg.get("upload", {}).get("delay_between_batches_seconds", 1)

    # Auto-discover Criteria Library Board and build name→item_id lookup
    criteria_map: dict = {}
    crit_lib_board_id = fetch_criteria_library_id(token, board_id)
    if crit_lib_board_id:
        log.info(f"Criteria Library Board ID: {crit_lib_board_id} — fetching items for relation linking...")
        criteria_map = fetch_criteria_library_map(token, crit_lib_board_id)
    else:
        log.warning("Could not auto-detect Criteria Library Board ID — board_relation column will be empty")

    # Create a new group named with today's date + dataset type
    today      = _date.today().strftime("%Y-%m-%d")
    group_name = f"{dataset_type.upper()}_{today}"

    if dry_run:
        log.info(f"[DRY-RUN] Would create group '{group_name}' on board {board_id}")
        group_id = "dry_run_group"
    else:
        log.info(f"Creating new group '{group_name}' on L3 board {board_id}...")
        group_id = create_group_on_board(token, board_id, group_name)
        if not group_id:
            log.error("Failed to create group on L3 board.")
            return {"created": 0, "skipped": 0, "errors": 1}
        log.info(f"  ✓ Group created: {group_id}")

    # Determine the item name column — "Name" (Deliverable name) is the Monday item title
    name_col = None
    for candidate in [
        "Name",
        "Criteria ID Number \u2013 Deliverable Name",
        "Deliverable VersionName",
        "Criteria ID Number",
        "Criteria Name",
    ]:
        if candidate in df.columns and df[candidate].notna().any():
            name_col = candidate
            break
    if not name_col:
        name_col = df.columns[0]
    log.info(f"Using '{name_col}' as L3 item name")

    # Log a sample date value so we can verify the format
    date_col = next((c for c in df.columns
                     if str(c).strip().lower() == "review date and timestamp"), None)
    if date_col:
        sample = df[date_col].dropna().head(1).values
        log.info(f"Review Date sample: {sample[0]!r}" if len(sample) else
                 "Review Date column found but all values are empty")
    else:
        log.warning(f"'Review Date and Timestamp' column not found in df. "
                    f"Available cols: {[c for c in df.columns if 'date' in str(c).lower() or 'review' in str(c).lower()]}")

    log.info(f"Uploading {len(df)} rows to L3 board (group: {group_name})...")
    stats = {"created": 0, "skipped": 0, "errors": 0}
    upload_rows = []

    for _, row in df.iterrows():
        item_name = str(row.get(name_col, "")).strip()
        if not item_name or item_name in ("nan", "<NA>", ""):
            stats["skipped"] += 1
            continue

        col_vals: dict = {}
        for pipeline_col, monday_col_id in resolved_cols.items():
            if monday_col_id == "name":
                continue  # item name is set separately

            val = row.get(pipeline_col)
            if val is None:
                continue
            try:
                if pd.isna(val):
                    continue
            except (TypeError, ValueError):
                pass
            val_str = str(val).strip()
            if val_str in ("", "nan", "<NA>", "NaT"):
                continue

            if monday_col_id == "date4":
                try:
                    date_part = pd.to_datetime(val_str).strftime("%Y-%m-%d")
                    col_vals[monday_col_id] = {"date": date_part}
                except Exception:
                    pass  # skip unparseable dates
            elif monday_col_id == "status":
                col_vals[monday_col_id] = {"label": val_str}
            elif monday_col_id == "numeric_mkw4vt0k":
                try:
                    col_vals[monday_col_id] = float(val_str)
                except (ValueError, TypeError):
                    pass
            elif monday_col_id == "long_text_mkw4w2ra":
                col_vals[monday_col_id] = {"text": val_str}
            else:
                col_vals[monday_col_id] = val_str

        # Link "Criteria (from Criteria Library Board)" board_relation using Criteria ID
        if criteria_map:
            crit_id_col = next((c for c in df.columns
                                if str(c).strip().lower() in (
                                    "criteria id",
                                    "criteria id number",
                                    "criteria id number \u2013 deliverable name",
                                )), None)
            crit_id_val = str(row.get(crit_id_col, "")).strip() if crit_id_col else ""
            # Try exact match first; also try with '_' swapped to '-' as a fallback
            crit_id_lookup = re.sub(r"([A-Z]{2,})[_\u2011\u2013](C\d+)", r"\1-\2", crit_id_val)
            linked_id = criteria_map.get(crit_id_val) or criteria_map.get(crit_id_lookup)
            if linked_id:
                col_vals["board_relation_mkvv6kmc"] = {"item_ids": [int(linked_id)]}
        upload_rows.append((item_name, col_vals))

    def _make_task(item_name: str, col_vals: dict):
        def _task():
            try:
                item_id = create_item(token, board_id, group_id, item_name, col_vals, dry_run)
                return {"ok": True, "item_id": item_id, "name": item_name}
            except Exception as e:
                return {"ok": False, "name": item_name, "error": str(e)}
        return _task

    tasks = [_make_task(item_name, col_vals) for item_name, col_vals in upload_rows]
    for result in _run_parallel_batches(tasks, batch_size=batch_size, delay=delay):
        if result["ok"]:
            if result.get("item_id"):
                log.info(f"  ✓ Created item {result['item_id']}: {result['name'][:65]!r}")
            stats["created"] += 1
        else:
            log.error(f"  ✗ Failed: {result['name']!r}: {result['error']}")
            stats["errors"] += 1

    stats["group_id"] = group_id   # expose so caller can trigger sync
    return stats


def fetch_group_item_ids(token: str, board_id: str, group_id: str) -> list:
    """
    Return a list of all item IDs within a specific group on a board.

    Used by sync_l3_group() to get the IDs of every item that was just uploaded,
    so the sync status column can be updated on each of them in turn.
    Limited to 500 items per the query; sufficient for typical upload batch sizes.

    Returns a list of item ID strings.
    """
    query = """
    query ($board_id: [ID!]!, $group_id: [String!]!) {
      boards(ids: $board_id) {
        groups(ids: $group_id) {
          items_page(limit: 500) { items { id } }
        }
      }
    }
    """
    data = _gql(token, query, {"board_id": board_id, "group_id": group_id})
    items = (data.get("boards", [{}])[0]
                 .get("groups", [{}])[0]
                 .get("items_page", {})
                 .get("items", []))
    return [item["id"] for item in items]


def fetch_status_labels(token: str, board_id: str, column_id: str) -> list:
    """
    Return all available label strings for a Monday.com status column.

    Status columns in Monday.com have a fixed set of labels defined in their
    settings (e.g. "Done", "In Progress", "Stuck"). The labels are stored as a
    dict inside the column's settings_str JSON field.

    Used by sync_l3_group() to validate the configured sync_label and to auto-
    select the first available label when none is configured.

    Returns a list of non-empty label strings, or an empty list if the column
    is not found or the settings JSON cannot be parsed.
    """
    query = """
    query ($board_id: [ID!]!) {
      boards(ids: $board_id) { columns { id settings_str } }
    }
    """
    data = _gql(token, query, {"board_id": board_id})
    for col in data.get("boards", [{}])[0].get("columns", []):
        if col.get("id") == column_id:
            try:
                settings = json.loads(col.get("settings_str", "{}"))
                labels = settings.get("labels", {})
                return [v for v in labels.values() if v]
            except (json.JSONDecodeError, AttributeError):
                pass
    return []


def sync_l3_group(token: str, board_id: str, group_id: str,
                  column_id: str = "color_mkwbaxzj",
                  label: Optional[str] = None,
                  dry_run: bool = False) -> dict:
    """
    Set the 'Synch Version and criteria' status column on every item in an L3 group.

    After uploading rows to the L3 board, a Monday.com automation needs to be
    triggered to link each item to its criteria version. The automation fires when
    the sync status column (color_mkwbaxzj) is set to a specific label.

    How it works:
    1. Fetch all available labels for the sync status column.
    2. Use the configured `label` from l3_sync.sync_label in config. If none is
       configured, auto-select the first available label (and warn). If no labels
       can be fetched at all, fall back to "Done".
    3. Warn if the configured label does not match any of the column's actual labels
       (a mismatch means the automation won't fire).
    4. Fetch all item IDs in the group and set the status on each one individually.

    The label parameter can be overridden at call time (e.g. from a POST request body
    sent by the app UI) without needing to edit the config file.

    Returns {"updated": int, "errors": int}.
    """
    labels = fetch_status_labels(token, board_id, column_id)
    log.info(f"Sync: available labels for '{column_id}': {labels}")

    if not label:
        if labels:
            label = labels[0]
            log.info(f"Sync: no label configured — using first available: '{label}'")
        else:
            label = "Done"
            log.warning("Sync: could not fetch labels — defaulting to 'Done'")
    else:
        if labels and label not in labels:
            log.warning(f"Sync: configured label '{label}' not in available labels {labels}")

    item_ids = fetch_group_item_ids(token, board_id, group_id)
    log.info(f"Sync: setting '{label}' on {len(item_ids)} items in group {group_id}...")

    mutation = """
    mutation ($board_id: ID!, $item_id: ID!, $col_id: String!, $val: JSON!) {
      change_column_value(board_id: $board_id, item_id: $item_id,
                          column_id: $col_id, value: $val) { id }
    }
    """
    stats = {"updated": 0, "errors": 0}
    col_val = json.dumps({"label": label})

    if dry_run:
        for item_id in item_ids:
            log.info(f"  [DRY-RUN] Would sync item {item_id}")
            stats["updated"] += 1
        log.info(f"Sync complete — updated: {stats['updated']}, errors: {stats['errors']}")
        return stats

    def _make_task(item_id: str):
        def _task():
            try:
                _gql(token, mutation, {
                    "board_id": board_id,
                    "item_id": item_id,
                    "col_id": column_id,
                    "val": col_val,
                })
                return {"ok": True, "item_id": item_id}
            except Exception as e:
                return {"ok": False, "item_id": item_id, "error": str(e)}
        return _task

    tasks = [_make_task(item_id) for item_id in item_ids]
    for result in _run_parallel_batches(tasks, batch_size=25, delay=0):
        if result["ok"]:
            stats["updated"] += 1
        else:
            log.error(f"  ✗ Sync failed for item {result['item_id']}: {result['error']}")
            stats["errors"] += 1

    log.info(f"Sync complete — updated: {stats['updated']}, errors: {stats['errors']}")
    return stats


# ──────────────────────────────────────────────────────────────
# CLI
# ──────────────────────────────────────────────────────────────

def list_all_boards(token: str) -> list:
    """
    Fetch up to 100 boards accessible to the given API token, ordered by creation date.

    Called by cmd_discover() as a fallback when the configured board_id is not found.
    Listing all boards helps the user identify the correct board ID and workspace
    when the token belongs to a different workspace than expected.

    Returns a list of board dicts, each containing id, name, and workspace info.
    """
    query = """
    query {
      boards(limit: 100, order_by: created_at) {
        id name workspace { id name }
      }
    }
    """
    data = _gql(token, query)
    return data.get("boards", [])


def cmd_discover(args):
    """
    CLI handler for the `discover` sub-command.

    Fetches and pretty-prints a board's group IDs and column IDs so the user can
    copy the right values into monday_config.json before running an upload.

    If the configured board_id is not accessible (wrong ID, wrong workspace, or
    insufficient token permissions), falls back to listing every board the token
    can see so the user can identify and correct the board ID.

    Accepts an optional --board flag to inspect any board by ID without changing
    the config file first.
    """
    cfg   = load_config()
    token = cfg["api_token"]
    bid   = args.board or cfg["board_id"]

    # First try the configured board ID
    log.info(f"Fetching board structure for board ID: {bid}")
    boards_raw = _gql(token, """
        query ($board_id: [ID!]!) {
          boards(ids: $board_id) {
            id name
            groups { id title }
            columns { id title type }
          }
        }
    """, {"board_id": [bid]}).get("boards", [])

    if not boards_raw:
        # Board not found — list all accessible boards so the user can find the right ID
        print(f"\n⚠️  Board ID '{bid}' was not found for this account.")
        print("   Your token may belong to a different workspace, or the board ID may be wrong.")
        print("\n🔍  Listing ALL boards accessible to your token...\n")
        try:
            all_boards = list_all_boards(token)
        except Exception as e:
            sys.exit(f"Could not list boards: {e}")

        if not all_boards:
            sys.exit(
                "No boards found. Make sure your API token has 'boards:read' permission.\n"
                "In Monday.com: go to a board → 3-dot menu → 'Get board link'.\n"
                "The number in the URL is the board ID. Update 'board_id' in monday_config.json."
            )

        print(f"  {'Board ID':<20}  {'Workspace':<25}  Board Name")
        print(f"  {'-'*20}  {'-'*25}  ----------")
        for b in all_boards:
            ws = b.get("workspace") or {}
            print(f"  {b['id']:<20}  {ws.get('name','(default)'):<25}  {b['name']}")

        print("\n💡  Copy the correct Board ID into monday_config.json → 'board_id'")
        print("    Then re-run:  python upload_to_monday.py --discover\n")
        return

    info = boards_raw[0]
    print("\n" + "═" * 65)
    print(f"  Board: {info['name']}  (id={info['id']})")
    print("═" * 65)

    print("\n📁  GROUPS  — copy IDs into monday_config.json → group_map:")
    print(f"  {'ID':<30}  Title")
    print(f"  {'-'*30}  -----")
    for g in info["groups"]:
        print(f"  {g['id']:<30}  {g['title']}")

    print("\n📋  COLUMNS  — copy IDs into monday_config.json → column_map:")
    print(f"  {'ID':<30}  {'Type':<20}  Title")
    print(f"  {'-'*30}  {'-'*20}  -----")
    for c in info["columns"]:
        print(f"  {c['id']:<30}  {c['type']:<20}  {c['title']}")

    print("\n💡  Fill in monday_config.json with the IDs above, then run:")
    print("    python upload_to_monday.py --upload --folder /path/to/data\n")


def cmd_upload(args):
    """
    CLI handler for the `upload` sub-command.

    Orchestrates the full upload sequence:
    1. Load data — either from a pre-processed .xlsx file (--file) or by running
       the pipeline on a folder of source files (--folder).
    2. Override dataset_type if --type was supplied explicitly.
    3. Upload unique Deliverable VersionNames to the L2 main board.
    4. Upload unique Criteria Names to the criteria-name-only board (if configured).
    5. Print a summary table with created / skipped / error counts for each board.

    In --dry-run mode all API calls are bypassed and only log messages are printed.
    """
    cfg = load_config()

    # Load data
    if args.file:
        df, dataset_type = load_from_file(args.file)
    elif args.folder:
        df, dataset_type = run_pipeline(args.folder, args.type or "auto")
    else:
        sys.exit("ERROR: Provide --folder or --file")

    if args.type and args.type != "auto":
        dataset_type = args.type

    if args.dry_run:
        log.info("DRY-RUN mode — no items will actually be created on Monday.com")

    log.info(f"Dataset type: {dataset_type.upper()}, rows: {len(df)}")

    # Upload unique Deliverable VersionNames to main board (FSD / CRD / PDD groups)
    log.info("\n── Uploading unique Deliverable VersionNames to main board ──")
    stats = upload_unique_names(df, dataset_type, cfg, dry_run=args.dry_run)

    # Upload unique Criteria Names to second board (if configured)
    log.info("\n── Uploading criteria names to criteria-name board ──")
    stats_crit = upload_criteria_names_only(df, cfg, dry_run=args.dry_run)

    print("\n" + "═" * 55)
    print("  Upload complete!")
    print(f"  L2 board (DVN items) — created: {stats['created']:>4}  "
          f"skipped: {stats['skipped']:>4}  errors: {stats['errors']:>4}")
    print(f"  Criteria board       — created: {stats_crit['created']:>4}  "
          f"skipped: {stats_crit['skipped']:>4}  errors: {stats_crit['errors']:>4}")
    print("═" * 55 + "\n")


def main():
    """
    Entry point for the command-line interface.

    Defines two sub-commands:
    - discover  Print board groups and column IDs to help populate monday_config.json.
    - upload    Run the pipeline (or load a file) and upload to Monday.com.

    Prints help and exits if no sub-command is provided.
    """
    parser = argparse.ArgumentParser(
        description="Upload processed Excel output to Monday.com.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python upload_to_monday.py discover\n"
            "  python upload_to_monday.py upload --file output.xlsx\n"
            "  python upload_to_monday.py upload --folder /path/to/data --dry-run\n"
        ),
    )
    sub = parser.add_subparsers(dest="command", metavar="command")

    # discover
    p_disc = sub.add_parser("discover", help="Print board groups and column IDs")
    p_disc.add_argument("--board", default=None, help="Override board ID (default: from config)")

    # upload
    p_up = sub.add_parser("upload", help="Upload unique Deliverable VersionNames to Monday.com")
    src  = p_up.add_mutually_exclusive_group(required=True)
    src.add_argument("--file",   metavar="PATH", help="Pre-processed output .xlsx file")
    src.add_argument("--folder", metavar="PATH", help="Folder of source Excel files (runs pipeline first)")
    p_up.add_argument("--type",    choices=["p2f", "qfs", "auto"], default="auto",
                      help="Force dataset type (default: auto-detect)")
    p_up.add_argument("--dry-run", action="store_true",
                      help="Simulate upload — print what would be created without actually sending")

    args = parser.parse_args()

    if args.command == "discover":
        cmd_discover(args)
    elif args.command == "upload":
        cmd_upload(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
