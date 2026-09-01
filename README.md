# Excel Pipeline — Setup & Usage Guide

A Flask web app that consolidates Excel quality-review files (P2F and Q&FS), produces a clean output, and uploads results to Monday.com boards (L2 and L3).

---

## 1. Prerequisites

- Python 3.9 or later
- Git
- A Monday.com account with API access

---

## 2. Clone the Repository

```bash
git clone https://github.ibm.com/Sara-Saad/Excel-Pipeline.git
cd Excel-Pipeline
```

---

## 3. Create a Virtual Environment

```bash
python3 -m venv ../venv
source ../venv/bin/activate
```

> On Windows use: `../venv/Scripts/activate`

---

## 4. Install Dependencies

```bash
pip install -r requirements.txt
```

---

## 5. Get Your Monday.com API Token

1. Log in to [monday.com](https://monday.com)
2. Click your **avatar** (bottom-left) → **Developers** → **API Tokens**
3. Click **Generate** (or copy your existing token)
4. Copy the full token string

---

## 6. Configure monday_config.json

Copy the example config and fill in your details:

```bash
cp monday_config.example.json monday_config.json
```

Open `monday_config.json` and replace the placeholders:

```json
{
  "api_token": "paste_your_token_here",

  "board_id": "18076098477",        ← L2 board ID (already set)

  "l3_board_id": "10065306343",     ← L3 board ID (already set)

  "l3_sync": {
    "column_id":  "color_mkwbaxzj",
    "sync_label": ""                
  }
}
```

> **Never commit `monday_config.json`** — it contains your API token and is listed in `.gitignore`.

### Find Board IDs (if you need to update them)

Run the discover command to list all boards your token can access:

```bash
python upload_to_monday.py discover
```

To inspect a specific board's groups and columns:

```bash
python upload_to_monday.py discover --board BOARD_ID
```

---

## 7. Run the App

```bash
python src/app.py
```

Open your browser at: **http://localhost:5000**

---

## 8. Using the App — Step by Step

### Step 1 — Load Your Excel Files

- Click **Choose Folder** and select the folder containing your `.xlsx` / `.xls` quality review files
- The app auto-detects whether the data is P2F or Q&FS format

### Step 2 — Run the Pipeline

- Click **Run Pipeline**
- The app consolidates all files, cleans the data, extracts Criteria IDs, strips timestamps from Review Dates, and produces a single output file

### Step 3 — Download the Output (optional, Please do as we need to manually check , with new files new corner cases will show up)

- Click **Download Excel** to save the consolidated output to your computer

### Step 4 — Upload to Monday L2

- Click **Upload to Monday L2**
- Uploads unique **Deliverable VersionNames** to the L2 board, routed to the correct group (FSD / CRD / PDD)

### Step 5 — Upload to Monday L3

- Click **Upload to Monday L3**
- Creates a new dated group on the L3 board (e.g. `P2F_2026-03-19`)
- Uploads every criteria row as one item with: Name, Criteria ID, Criteria Name, Review Date, Status, Score, Deliverable VersionName, Detailed Comments
- Automatically links the **Criteria (from Criteria Library Board)** board relation using the Criteria ID

### Step 6 — Sync L3

- After L3 upload completes, the **Sync L3** button activates
- Click it to set the **Synch Version and criteria** status on all newly uploaded rows, which triggers the Monday.com automation



## 9. Project Structure

```
Excel-Pipeline/
├── src/
│   ├── app.py              # Flask web server + UI + Monday.com routes
│   ├── combiner.py         # Loads and aligns Excel files into a DataFrame
│   ├── data_q2.py          # Core data quality and schema transformation
│   ├── processor.py        # File-type detection and preprocessing
│   ├── pipeline.py         # Orchestrates the full processing pipeline
│   ├── checker.py          # Column validation helpers
│   └── __init__.py
├── upload_to_monday.py     # Monday.com API integration (L2, L3, Sync)
├── monday_config.example.json  # Config template (safe to commit)
├── monday_config.json      # Your config with real token (gitignored)
├── requirements.txt        # Python dependencies
└── README.md
```

---

## 10. Running the Monday.com CLI Directly (optional)

You can also trigger uploads from the command line without the web UI:

```bash
# Discover board structure
python upload_to_monday.py discover
python upload_to_monday.py discover --board 10065306343

# Upload from a pre-processed output file
python upload_to_monday.py upload --file /path/to/output.xlsx

# Upload from a folder of source files (runs pipeline first)
python upload_to_monday.py upload --folder /path/to/data

# Dry run — prints what would be uploaded without sending anything
python upload_to_monday.py upload --file output.xlsx --dry-run
```
