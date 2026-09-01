from flask import Flask, request, render_template_string, send_file, jsonify
import pandas as pd
import os
import sys
import glob
import io
import json
import traceback
from datetime import datetime
import re
import threading

# Ensure the Excel-Pipeline root (parent of src/) is on sys.path so that
# 'from src.combiner import ...' works whether the app is launched as:
#   python src/app.py          (from Excel-Pipeline/)
#   gunicorn src.app:app       (from Excel-Pipeline/)
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

app = Flask(__name__)

# Use the consolidated logic from src/
from src.combiner import combine_folder_to_frames
from src.data_q2 import to_p2f_final_schema, to_qfs_final_schema

# ── Monday.com auto-upload (optional — only runs if monday_config.json is configured) ──
def _trigger_monday_upload(df: pd.DataFrame, dataset_type: str) -> None:
    """
    Fire-and-forget Monday.com upload after the pipeline completes.
    Runs in a background thread so it never blocks the HTTP response.
    """
    try:
        config_path = os.path.join(_ROOT, "monday_config.json")
        if not os.path.exists(config_path):
            return
        with open(config_path) as f:
            cfg = json.load(f)
        token = cfg.get("api_token", "")
        if not token or token == "YOUR_API_TOKEN_HERE":
            return  # Not yet configured — silently skip

        # Import upload helpers from the Excel-Pipeline root
        import importlib.util, sys as _sys
        spec = importlib.util.spec_from_file_location(
            "upload_to_monday",
            os.path.join(_ROOT, "upload_to_monday.py")
        )
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)

        app.logger.info("monday · starting background upload")
        s1 = mod.upload_unique_names(df, dataset_type, cfg, dry_run=False)
        s2 = mod.upload_criteria_names_only(df, cfg, dry_run=False)
        app.logger.info(
            f"monday · done — L2 board: +{s1['created']} created, "
            f"{s1['skipped']} skipped, {s1['errors']} errors | "
            f"criteria board: +{s2['created']} created"
        )
    except Exception as _e:
        app.logger.warning(f"monday · upload failed (non-fatal): {_e}")

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>IBM — Enablement Team Orchestrator</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@300;400;500&display=swap" rel="stylesheet">
    <style>
        :root {
            /* IBM Carbon (light) inspired tokens */
            --bg:            #f4f4f4;
            --layer-01:      #ffffff;
            --layer-02:      #f4f4f4;
            --layer-03:      #e0e0e0;
            --border-subtle: #e0e0e0;
            --border-strong: #8d8d8d;

            --text:          #161616;
            --text-secondary:#525252;
            --text-muted:    #6f6f6f;

            --focus:         #0f62fe;
            --interactive:   #0f62fe;
            --interactive-h: #0353e9;

            --support-success:#24a148;
            --support-error:  #da1e28;
            --support-warn:   #f1c21b;
        }

        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            background: var(--bg);
            color: var(--text);
            font-family: 'IBM Plex Sans', sans-serif;
            min-height: 100vh;
            overflow-x: hidden;
        }

        /* Clean background (no grid) */
        body::before,
        body::after {
            content: none;
        }

        /* Carbon-like masthead */
        .masthead {
            position: sticky;
            top: 0;
            z-index: 10;
            background: var(--layer-01);
            border-bottom: 1px solid var(--border-subtle);
        }

        .masthead-inner {
            max-width: 1400px;
            margin: 0 auto;
            padding: 14px 48px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 16px;
        }

        .brand {
            display: flex;
            align-items: center;
            gap: 14px;
            min-width: 0;
        }

        .brand-title {
            display: flex;
            flex-direction: column;
            gap: 2px;
            min-width: 0;
        }

        .brand-title .product {
            font-size: 14px;
            font-weight: 600;
            letter-spacing: 0.2px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .brand-title .desc {
            font-family: 'IBM Plex Mono', monospace;
            font-size: 11px;
            color: var(--text-muted);
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .container {
            position: relative;
            z-index: 1;
            max-width: 1400px;
            margin: 0 auto;
            padding: 28px 48px 60px;
        }

        @media (max-width: 900px) {
            .container { padding: 20px 20px 40px; }
        }

        /* Header */
        .header {
            margin-bottom: 24px;
            animation: fadeDown 0.6s ease both;
        }

        .logo {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 24px;
        }

        h1 {
            font-size: clamp(28px, 4.2vw, 44px);
            font-weight: 700;
            letter-spacing: -1.5px;
            line-height: 1.06;
            margin-bottom: 16px;
        }

        h1 em {
            font-style: normal;
            background: linear-gradient(90deg, var(--interactive), #78a9ff);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
        }

        .subtitle {
            font-size: 16px;
            color: var(--text-muted);
            font-family: 'IBM Plex Mono', monospace;
            font-weight: 300;
        }

        /* Pipeline visual */
        .pipeline-steps {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 48px;
            padding: 20px 24px;
            background: var(--layer-01);
            border: 1px solid var(--border-subtle);
            border-radius: 10px;
            overflow-x: auto;
            animation: fadeUp 0.6s 0.1s ease both;
        }

        .pipe-step {
            display: flex;
            align-items: center;
            gap: 8px;
            flex-shrink: 0;
        }

        .pipe-badge {
            padding: 6px 14px;
            border-radius: 0;
            font-size: 12px;
            font-weight: 600;
            font-family: 'IBM Plex Mono', monospace;
            border: 1px solid;
            transition: all 0.3s;
            letter-spacing: 0.5px;
            text-transform: uppercase;
        }

        .pipe-badge.idle {
            background: rgba(255,255,255,0.03);
            border-color: var(--border-subtle);
            color: var(--text-muted);
        }

        .pipe-badge.running {
            background: rgba(15,98,254,0.15);
            border-color: var(--interactive);
            color: #78a9ff;
            animation: pulse 1s ease-in-out infinite;
        }

        .pipe-badge.done {
            background: rgba(66,190,101,0.12);
            border-color: var(--support-success);
            color: var(--support-success);
        }

        .pipe-badge.error {
            background: rgba(250,77,86,0.12);
            border-color: var(--support-error);
            color: var(--support-error);
        }

        .pipe-arrow {
            color: var(--border-strong);
            font-size: 18px;
            font-weight: 300;
        }

        /* Input section */
        .input-card {
            background: var(--layer-01);
            border: 1px solid var(--border-subtle);
            border-radius: 12px;
            padding: 28px;
            margin-bottom: 32px;
            animation: fadeUp 0.6s 0.2s ease both;
            transition: border-color 0.3s;
            box-shadow: 0 8px 24px rgba(0,0,0,0.06);
        }

        .input-card:hover {
            border-color: rgba(15,98,254,0.5);
        }

        .input-label {
            display: block;
            font-size: 13px;
            font-weight: 600;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            color: var(--text-muted);
            margin-bottom: 12px;
        }

        .input-row {
            display: flex;
            gap: 12px;
        }

        .path-input {
            flex: 1;
            background: var(--layer-02);
            border: 1px solid var(--border-subtle);
            border-radius: 10px;
            padding: 14px 16px;
            font-family: 'IBM Plex Mono', monospace;
            font-size: 14px;
            color: var(--text);
            outline: none;
            transition: all 0.2s;
        }

        .path-input:focus {
            border-color: var(--focus);
            box-shadow: 0 0 0 2px rgba(15,98,254,0.35);
        }

        .path-input::placeholder {
            color: var(--text-muted);
        }

        .run-btn {
            padding: 14px 22px;
            background: var(--interactive);
            border: none;
            border-radius: 0;
            color: white;
            font-family: 'IBM Plex Sans', sans-serif;
            font-size: 15px;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.2s;
            white-space: nowrap;
            position: relative;
            overflow: hidden;
            min-width: 160px;
            border-radius: 10px;
        }

        .run-btn::before {
            content: '';
            position: absolute;
            inset: 0;
            background: linear-gradient(135deg, rgba(255,255,255,0.15), transparent);
            opacity: 0;
            transition: opacity 0.2s;
        }

        .run-btn:hover:not(:disabled)::before {
            opacity: 1;
        }

        .run-btn:hover:not(:disabled) {
            background: var(--interactive-h);
            box-shadow: 0 6px 20px rgba(0,0,0,0.35);
        }

        .run-btn:active:not(:disabled) {
            transform: translateY(0);
        }

        .run-btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .input-hint {
            margin-top: 10px;
            font-family: 'IBM Plex Mono', monospace;
            font-size: 12px;
            color: var(--text-muted);
        }

        /* Log section */
        .log-card {
            background: var(--layer-01);
            border: 1px solid var(--border-subtle);
            border-radius: 12px;
            overflow: hidden;
            animation: fadeUp 0.6s 0.3s ease both;
            display: none;
            box-shadow: 0 8px 24px rgba(0,0,0,0.06);
        }

        .log-card.visible {
            display: block;
        }

        .log-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 20px 28px;
            border-bottom: 1px solid var(--border-subtle);
            background: rgba(255,255,255,0.02);
        }

        .log-title {
            font-size: 13px;
            font-weight: 700;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            color: var(--text-muted);
            display: flex;
            align-items: center;
            gap: 8px;
        }

        .live-dot {
            width: 7px;
            height: 7px;
            background: var(--support-success);
            border-radius: 50%;
            animation: blink 1s ease-in-out infinite;
        }

        .log-body {
            padding: 24px 28px;
            font-family: 'IBM Plex Mono', monospace;
            font-size: 13px;
            line-height: 1.8;
            max-height: 360px;
            overflow-y: auto;
        }

        .log-body::-webkit-scrollbar {
            width: 4px;
        }

        .log-body::-webkit-scrollbar-track {
            background: transparent;
        }

        .log-body::-webkit-scrollbar-thumb {
            background: var(--border-strong);
            border-radius: 2px;
        }

        .log-line {
            display: flex;
            gap: 16px;
            opacity: 0;
            animation: logIn 0.3s ease forwards;
        }

        .log-time {
            color: var(--muted);
            flex-shrink: 0;
        }

        .log-msg { flex: 1; }
        .log-msg.success { color: var(--support-success); }
        .log-msg.error { color: var(--support-error); }
        .log-msg.info { color: #78a9ff; }
        .log-msg.warn { color: var(--support-warn); }
        .log-msg.dim { color: var(--text-muted); }

        /* Results */
        .result-card {
            background: var(--layer-01);
            border: 1px solid var(--border-subtle);
            border-radius: 12px;
            padding: 28px;
            margin-top: 24px;
            display: none;
            box-shadow: 0 8px 24px rgba(0,0,0,0.06);
        }

        .result-card.visible {
            display: block;
            animation: fadeUp 0.4s ease both;
        }

        .result-card.success-card {
            border-color: rgba(66,190,101,0.35);
            background: linear-gradient(135deg, rgba(66,190,101,0.05), var(--layer-01));
        }

        .result-card.error-card {
            border-color: rgba(250,77,86,0.35);
            background: linear-gradient(135deg, rgba(250,77,86,0.05), var(--layer-01));
        }

        .result-header {
            display: flex;
            align-items: center;
            gap: 14px;
            margin-bottom: 20px;
        }

        .result-icon {
            width: 44px;
            height: 44px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 20px;
            flex-shrink: 0;
        }

        .result-icon.ok { background: rgba(74,222,128,0.15); }
        .result-icon.err { background: rgba(250,77,86,0.15); }

        .result-title {
            font-size: 20px;
            font-weight: 700;
        }

        .result-meta {
            font-size: 13px;
            color: var(--muted);
            font-family: 'IBM Plex Mono', monospace;
            margin-top: 2px;
        }

        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
            gap: 12px;
            margin-bottom: 24px;
        }

        .stat-item {
            background: var(--layer-02);
            border: 1px solid var(--border-subtle);
            border-radius: 12px;
            padding: 16px;
        }

        .stat-value {
            font-size: 26px;
            font-weight: 800;
            color: var(--support-success);
            letter-spacing: -1px;
        }

        .stat-label {
            font-size: 11px;
            color: var(--text-muted);
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-top: 4px;
            font-family: 'IBM Plex Mono', monospace;
        }

        .action-row {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            margin-top: 8px;
        }
        .action-row a,
        .action-row button,
        .action-row label {
            flex: 1 1 200px;
            text-align: center;
            box-sizing: border-box;
        }
        .action-row-full {
            flex: 1 1 100%;
        }

        .download-btn {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            padding: 14px 24px;
            background: var(--support-success);
            border: none;
            border-radius: 10px;
            color: #ffffff;
            font-family: 'IBM Plex Sans', sans-serif;
            font-size: 14px;
            font-weight: 700;
            cursor: pointer;
            text-decoration: none;
            transition: all 0.2s;
            white-space: nowrap;
        }

        .download-btn:hover {
            filter: brightness(0.95);
            box-shadow: 0 6px 18px rgba(0,0,0,0.35);
        }

        .monday-btn {
            display: inline-flex;
            align-items: center;
            gap: 8px;
            padding: 14px 24px;
            background: #f6f7fb;
            border: 2px solid #c3c6d4;
            border-radius: 10px;
            color: #323338;
            font-family: 'IBM Plex Sans', sans-serif;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.2s;
            white-space: nowrap;
        }

        .monday-btn:hover:not(:disabled) {
            background: #e6e9f4;
            border-color: #6161ff;
            color: #6161ff;
            box-shadow: 0 4px 14px rgba(97,97,255,0.18);
        }

        .monday-btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        /* ── In-app editor (full-screen) ── */
        .editor-overlay {
            display: none;
            position: fixed; inset: 0;
            background: var(--layer-01);
            z-index: 1000;
            flex-direction: column;
        }
        .editor-overlay.open { display: flex; }

        .editor-topbar {
            display: flex; align-items: center; gap: 16px;
            padding: 0 24px;
            height: 56px;
            background: var(--ui-bg);
            border-bottom: 2px solid var(--focus);
            flex-shrink: 0;
        }
        .editor-topbar-title {
            font-size: 16px; font-weight: 600;
            color: var(--text-primary); flex: 1;
            letter-spacing: 0.01em;
        }
        .editor-topbar-meta {
            font-size: 12px; color: var(--text-secondary);
            font-family: 'IBM Plex Mono', monospace;
        }
        .editor-save-btn {
            height: 40px; padding: 0 24px;
            background: var(--focus); color: #fff;
            border: none; border-radius: 2px;
            font-size: 14px; font-weight: 600;
            cursor: pointer; letter-spacing: 0.02em;
            transition: background 0.15s;
        }
        .editor-save-btn:hover { background: #0043a8; }
        .editor-close-btn {
            height: 40px; padding: 0 20px;
            background: transparent; color: var(--text-secondary);
            border: 1px solid var(--border-subtle); border-radius: 2px;
            font-size: 14px; cursor: pointer;
            transition: background 0.15s, color 0.15s;
        }
        .editor-close-btn:hover { background: var(--layer-02); color: var(--text-primary); }

        /* action toolbar */
        .editor-toolbar2 {
            display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
            padding: 7px 20px;
            background: var(--layer-02);
            border-bottom: 1px solid var(--border-subtle);
            flex-shrink: 0;
        }
        .editor-global-search {
            height: 32px; width: 260px;
            padding: 0 10px;
            background: var(--layer-01);
            border: 1px solid var(--border-subtle); border-radius: 2px;
            color: var(--text-primary); font-size: 13px;
            font-family: 'IBM Plex Sans', sans-serif;
            outline: none; flex-shrink: 0;
        }
        .editor-global-search:focus { border-color: var(--focus); }
        .editor-tb-divider {
            width: 1px; height: 24px;
            background: var(--border-subtle); flex-shrink: 0;
            margin: 0 4px;
        }
        .editor-tb-btn {
            height: 30px; padding: 0 14px;
            background: var(--layer-01);
            border: 1px solid var(--border-subtle); border-radius: 2px;
            color: var(--text-primary); font-size: 12px; font-weight: 500;
            cursor: pointer; white-space: nowrap;
            transition: background 0.12s;
        }
        .editor-tb-btn:hover:not(:disabled) { background: var(--layer-03); }
        .editor-tb-btn:disabled { opacity: 0.4; cursor: not-allowed; }
        .editor-tb-btn.danger { color: var(--support-error); border-color: var(--support-error); }
        .editor-tb-btn.danger:hover:not(:disabled) { background: rgba(218,30,40,0.08); }
        .editor-row-count, .editor-sel-count {
            font-size: 12px; color: var(--text-secondary);
            font-family: 'IBM Plex Mono', monospace;
            margin-left: 4px;
        }
        .editor-sel-count { color: var(--focus); font-weight: 600; }

        .editor-del-btn {
            padding: 2px 7px; font-size: 11px;
            background: transparent; color: var(--support-error);
            border: 1px solid transparent; border-radius: 2px;
            cursor: pointer; opacity: 0.5;
            transition: opacity 0.1s, border-color 0.1s;
        }
        .editor-del-btn:hover { opacity: 1; border-color: var(--support-error); }

        .editor-table-wrap {
            overflow: auto; flex: 1;
        }
        .editor-table {
            border-collapse: collapse;
            font-size: 13px;
            font-family: 'IBM Plex Sans', sans-serif;
            width: 100%;
            table-layout: auto;
        }
        .editor-table thead tr {
            position: sticky; top: 0; z-index: 3;
        }
        .editor-table th {
            background: var(--layer-02);
            border-right: 1px solid var(--border-subtle);
            border-bottom: 2px solid var(--border-subtle);
            padding: 9px 12px;
            text-align: left; font-weight: 600;
            font-size: 11px; white-space: nowrap;
            color: var(--text-secondary);
            text-transform: uppercase; letter-spacing: 0.06em;
        }
        .editor-table th:first-child {
            width: 40px; text-align: center;
            background: var(--layer-03);
        }
        .editor-table tr:nth-child(even) td { background: var(--layer-02); }
        .editor-table tr:hover td { background: #e5efff !important; }
        .editor-table td {
            border-right: 1px solid var(--border-subtle);
            border-bottom: 1px solid var(--border-subtle);
            padding: 0;
            min-width: 100px;
        }
        .editor-table td.del-cell {
            width: 40px; min-width: 40px; max-width: 40px;
            text-align: center; padding: 4px;
            background: var(--layer-02) !important;
        }
        .editor-table td[contenteditable] {
            padding: 7px 12px;
            outline: none;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            max-width: 320px;
            cursor: text;
        }
        .editor-table td[contenteditable]:focus {
            background: #dbeeff !important;
            outline: 2px solid var(--focus);
            outline-offset: -2px;
            white-space: pre-wrap;
            overflow: visible;
        }
        .editor-table tr.hidden-row { display: none; }

        .editor-statusbar {
            display: flex; align-items: center; gap: 16px;
            padding: 0 24px;
            height: 36px;
            background: var(--layer-02);
            border-top: 1px solid var(--border-subtle);
            flex-shrink: 0;
            font-size: 12px;
        }
        .editor-statusbar-msg { flex: 1; font-family: 'IBM Plex Mono', monospace; }
        .editor-statusbar-msg.ok  { color: var(--support-success); }
        .editor-statusbar-msg.err { color: var(--support-error); }
        .editor-statusbar-hint { color: var(--text-secondary); font-size: 11px; }

        .monday-btn .monday-logo {
            width: 18px;
            height: 18px;
        }

        .monday-status {
            font-size: 12px;
            font-family: 'IBM Plex Mono', monospace;
            margin-top: 6px;
            padding: 6px 10px;
            border-radius: 6px;
            display: none;
        }

        .monday-status.ok  { background: rgba(36,161,72,0.1);  color: #24a148; display:block; }
        .monday-status.err { background: rgba(218,30,40,0.1);  color: #da1e28; display:block; }
        .monday-status.pending { background: rgba(15,98,254,0.1); color: #0f62fe; display:block; }

        /* Spinner */
        .spinner {
            display: inline-block;
            width: 14px;
            height: 14px;
            border: 2px solid rgba(0,0,0,0.18);
            border-top-color: rgba(0,0,0,0.6);
            border-radius: 50%;
            animation: spin 0.6s linear infinite;
            vertical-align: middle;
            margin-right: 6px;
        }

        /* Animations */
        @keyframes fadeDown {
            from { opacity: 0; transform: translateY(-20px); }
            to { opacity: 1; transform: translateY(0); }
        }

        @keyframes fadeUp {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }

        @keyframes logIn {
            from { opacity: 0; transform: translateX(-8px); }
            to { opacity: 1; transform: translateX(0); }
        }

        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.6; }
        }

        @keyframes blink {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.2; }
        }

        @keyframes spin {
            to { transform: rotate(360deg); }
        }

        /* Error message styling */
        .error-text {
            font-family: 'IBM Plex Mono', monospace;
            font-size: 13px;
            color: var(--support-error);
            background: rgba(250,77,86,0.06);
            border: 1px solid rgba(250,77,86,0.25);
            border-radius: 12px;
            padding: 16px;
            line-height: 1.7;
            white-space: pre-wrap;
        }

        /* ── Masthead extras ── */
        .brand-sep {
            width: 1px; height: 28px;
            background: var(--border-subtle);
            margin: 0 16px;
            flex-shrink: 0;
        }

        .masthead-status {
            display: flex; align-items: center; gap: 7px;
            font-size: 12px; color: var(--text-muted);
            font-family: 'IBM Plex Mono', monospace;
        }
        .status-dot {
            width: 7px; height: 7px; border-radius: 50%;
            background: var(--support-success);
            animation: blink 2.4s ease-in-out infinite;
            flex-shrink: 0;
        }

        /* ── Capabilities strip ── */
        .caps-strip {
            display: flex; align-items: stretch;
            gap: 0;
            background: var(--layer-01);
            border: 1px solid var(--border-subtle);
            border-radius: 12px;
            margin-bottom: 28px;
            overflow: hidden;
            box-shadow: 0 4px 16px rgba(0,0,0,0.05);
            animation: fadeUp 0.6s 0.15s ease both;
        }

        .cap-card {
            flex: 1; display: flex; align-items: flex-start; gap: 14px;
            padding: 22px 24px;
            transition: background 0.2s;
        }
        .cap-card:hover { background: rgba(15,98,254,0.03); }

        .cap-icon {
            width: 36px; height: 36px; flex-shrink: 0;
            display: flex; align-items: center; justify-content: center;
            border-radius: 8px;
            background: rgba(15,98,254,0.08);
            color: var(--interactive);
            font-size: 17px;
            font-weight: 700;
            line-height: 1;
        }

        .cap-body { display: flex; flex-direction: column; gap: 4px; }
        .cap-title {
            font-size: 13px; font-weight: 700;
            letter-spacing: 0.5px; text-transform: uppercase;
            color: var(--text);
        }
        .cap-desc {
            font-size: 13px; line-height: 1.6;
            color: var(--text-muted);
        }

        .cap-divider {
            width: 1px; background: var(--border-subtle);
            align-self: stretch; flex-shrink: 0;
            margin: 16px 0;
        }

        @media (max-width: 760px) {
            .caps-strip { flex-direction: column; }
            .cap-divider { width: auto; height: 1px; margin: 0 16px; }
        }

        /* Reduce motion a bit */
        @media (prefers-reduced-motion: reduce) {
            * { animation: none !important; transition: none !important; }
        }

    </style>
</head>
<body>
    <div class="masthead">
        <div class="masthead-inner">
            <div class="brand">
                <img src="https://upload.wikimedia.org/wikipedia/commons/5/51/IBM_logo.svg" alt="IBM" height="30" style="display:block;">
                <div class="brand-sep"></div>
                <div class="brand-title">
                    <span class="product">Enablement Team Orchestrator</span>
                    <span class="desc">Data Processing &amp; Monday.com Integration</span>
                </div>
            </div>
            <div class="masthead-status">
                <span class="status-dot"></span>
                <span class="status-label">Ready</span>
            </div>
        </div>
    </div>
    <div class="container">
        <!-- Header -->
        <div class="header">
            <h1>Enablement Team<br><em>Orchestrator</em></h1>
            <p class="subtitle">Combine, validate, and publish Excel data to Monday.com in a single run.</p>
        </div>

        <!-- Capabilities strip -->
        <div class="caps-strip">
            <div class="cap-card">
                <div class="cap-icon">⚙</div>
                <div class="cap-body">
                    <div class="cap-title">Process</div>
                    <div class="cap-desc">Merges all source Excel files in a folder, deduplicates columns, and applies the P2F or Q&FS schema automatically.</div>
                </div>
            </div>
            <div class="cap-divider"></div>
            <div class="cap-card">
                <div class="cap-icon">✦</div>
                <div class="cap-body">
                    <div class="cap-title">Validate</div>
                    <div class="cap-desc">Runs quality checks, enforces column types, resolves naming inconsistencies, and flags missing or malformed values.</div>
                </div>
            </div>
            <div class="cap-divider"></div>
            <div class="cap-card">
                <div class="cap-icon">↑</div>
                <div class="cap-body">
                    <div class="cap-title">Publish</div>
                    <div class="cap-desc">Uploads unique Deliverable Versions to the L2 board and detailed criteria rows to L3 — then triggers the sync automation.</div>
                </div>
            </div>
        </div>

        <!-- Input -->
        <div class="input-card">
            <label class="input-label">Source folder</label>
            <div class="input-row">
                <input 
                    class="path-input" 
                    type="text" 
                    id="folderPath"
                    placeholder="/path/to/your/excel/files"
                >
                <button class="run-btn" id="runBtn" onclick="runPipeline()">
                    Run Pipeline
                </button>
            </div>
            <div class="input-hint">Supports <code style="font-family:'IBM Plex Mono',monospace;font-size:11px;background:var(--layer-02);border:1px solid var(--border-subtle);border-radius:3px;padding:1px 6px;color:var(--focus);">.xlsx</code> · Auto-detects P2F and Q&FS file types · Outputs a single combined workbook</div>
        </div>

        <!-- Log output -->
        <div class="log-card" id="logCard">
            <div class="log-header">
                <div class="log-title">
                    <div class="live-dot"></div>
                    Pipeline Log
                </div>
                <span style="font-family:'IBM Plex Mono',monospace;font-size:12px;color:var(--muted)" id="logCount">0 events</span>
            </div>
            <div class="log-body" id="logBody"></div>
        </div>

        <!-- Result -->
        <div class="result-card" id="resultCard">
            <div class="result-header">
                <div class="result-icon" id="resultIcon"></div>
                <div>
                    <div class="result-title" id="resultTitle"></div>
                    <div class="result-meta" id="resultMeta"></div>
                </div>
            </div>
            <div id="resultContent"></div>
        </div>
    </div>

    <script>
        let logCount = 0;

        function addLog(msg, type = '') {
            const body = document.getElementById('logBody');
            const time = new Date().toLocaleTimeString('en', { hour12: false });
            const line = document.createElement('div');
            line.className = 'log-line';
            line.style.animationDelay = (logCount * 0.05) + 's';
            line.innerHTML = `
                <span class="log-time">${time}</span>
                <span class="log-msg ${type}">${msg}</span>
            `;
            body.appendChild(line);
            body.scrollTop = body.scrollHeight;
            logCount++;
            document.getElementById('logCount').textContent = logCount + ' events';
        }

        function setStep(step, state) {
            const el = document.getElementById('step-' + step);
            if (el) {
                el.className = 'pipe-badge ' + state;
            }
        }

        function resetAll() {
            logCount = 0;
            document.getElementById('logBody').innerHTML = '';
            document.getElementById('logCount').textContent = '0 events';
            document.getElementById('resultCard').className = 'result-card';
            ['processor', 'quality', 'checker', 'combiner', 'export'].forEach(s => setStep(s, 'idle'));
        }

        async function runPipeline() {
            const folderPath = document.getElementById('folderPath').value.trim();
            if (!folderPath) {
                document.getElementById('folderPath').focus();
                return;
            }

            resetAll();

            const btn = document.getElementById('runBtn');
            btn.disabled = true;
            btn.innerHTML = '<span class="spinner"></span>Processing…';

            document.getElementById('logCard').className = 'log-card visible';

            addLog('Starting pipeline…', 'info');
            addLog('Source: ' + folderPath, 'dim');

            try {
                const response = await fetch('/run_pipeline', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ folder_path: folderPath })
                });

                const data = await response.json();

                // Show log lines from server
                if (data.logs) {
                    data.logs.forEach(entry => {
                        addLog(entry.msg, entry.type || '');
                        if (entry.step) setStep(entry.step, entry.step_state || 'done');
                    });
                }

                if (data.success) {
                    setStep('export', 'done');
                    showSuccess(data);
                } else {
                    setStep(data.failed_step || 'export', 'error');
                    showError(data.error || 'Pipeline failed');
                }
            } catch (err) {
                addLog('Network error: ' + err.message, 'error');
                showError('Could not connect to server.');
            }

            btn.disabled = false;
            btn.innerHTML = 'Run Pipeline';
        }

        function showSuccess(data) {
            const card = document.getElementById('resultCard');
            card.className = 'result-card visible success-card';
            document.getElementById('resultIcon').className = 'result-icon ok';
            document.getElementById('resultIcon').textContent = '✓';
            document.getElementById('resultTitle').textContent = 'Pipeline complete';
            document.getElementById('resultMeta').textContent = 'All 5 steps executed successfully';

            const statsHtml = `
                <div class="stats-grid">
                    <div class="stat-item">
                        <div class="stat-value">${data.files_processed || 0}</div>
                        <div class="stat-label">Files Processed</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-value">${data.total_rows || 0}</div>
                        <div class="stat-label">Total Rows</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-value">${data.total_cols || 0}</div>
                        <div class="stat-label">Columns</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-value">${data.duration || '–'}</div>
                        <div class="stat-label">Duration (s)</div>
                    </div>
                </div>

                <div style="margin:20px 0 8px;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--text-secondary);">Export</div>
                <div class="action-row">
                    <a class="download-btn" href="/download" download="combined_output.xlsx">
                        ↓ &nbsp; Download Excel
                    </a>
                    <button class="download-btn" style="cursor:pointer;" onclick="openEditor()">
                        ✎ &nbsp; Inspect &amp; Edit
                    </button>
                    <label class="download-btn action-row-full" style="cursor:pointer;">
                        ↑ &nbsp; Upload Edited Excel
                        <input type="file" id="editedFileInput" accept=".xlsx,.xls"
                               style="display:none" onchange="uploadEdited(this)">
                    </label>
                </div>
                <div class="monday-status" id="editedUploadStatus"></div>

                <div style="margin:20px 0 8px;font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:var(--text-secondary);">Monday.com</div>
                <div class="action-row">
                    <button class="monday-btn" id="mondayL2Btn" onclick="uploadToMonday('l2')">
                        <svg class="monday-logo" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <circle cx="7"  cy="30" r="6" fill="#ff3d57"/>
                            <circle cx="20" cy="30" r="6" fill="#ffcc00"/>
                            <circle cx="33" cy="30" r="6" fill="#00ca72"/>
                            <path d="M4 24 L10 12" stroke="#ff3d57" stroke-width="3.5" stroke-linecap="round"/>
                            <path d="M17 24 L23 12" stroke="#ffcc00" stroke-width="3.5" stroke-linecap="round"/>
                            <path d="M30 24 L36 12" stroke="#00ca72" stroke-width="3.5" stroke-linecap="round"/>
                        </svg>
                        Upload to Monday L2
                    </button>
                    <button class="monday-btn" id="mondayL3Btn" onclick="uploadToMonday('l3')">
                        <svg class="monday-logo" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <circle cx="7"  cy="30" r="6" fill="#ff3d57"/>
                            <circle cx="20" cy="30" r="6" fill="#ffcc00"/>
                            <circle cx="33" cy="30" r="6" fill="#00ca72"/>
                            <path d="M4 24 L10 12" stroke="#ff3d57" stroke-width="3.5" stroke-linecap="round"/>
                            <path d="M17 24 L23 12" stroke="#ffcc00" stroke-width="3.5" stroke-linecap="round"/>
                            <path d="M30 24 L36 12" stroke="#00ca72" stroke-width="3.5" stroke-linecap="round"/>
                        </svg>
                        Upload to Monday L3
                    </button>
                    <button class="monday-btn action-row-full" id="mondaySyncL3Btn" onclick="syncL3()" disabled>
                        <svg class="monday-logo" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
                            <circle cx="7"  cy="30" r="6" fill="#ff3d57"/>
                            <circle cx="20" cy="30" r="6" fill="#ffcc00"/>
                            <circle cx="33" cy="30" r="6" fill="#00ca72"/>
                            <path d="M4 24 L10 12" stroke="#ff3d57" stroke-width="3.5" stroke-linecap="round"/>
                            <path d="M17 24 L23 12" stroke="#ffcc00" stroke-width="3.5" stroke-linecap="round"/>
                            <path d="M30 24 L36 12" stroke="#00ca72" stroke-width="3.5" stroke-linecap="round"/>
                        </svg>
                        Sync L3
                    </button>
                </div>
                <div class="monday-status" id="mondayStatusL2"></div>
                <div class="monday-status" id="mondayStatusL3"></div>
                <div class="monday-status" id="mondayStatusSyncL3"></div>
            `;
            document.getElementById('resultContent').innerHTML = statsHtml;
        }

        async function uploadToMonday(level) {
            // level is 'l2' or 'l1' — extract the number to build correct element IDs
            const num      = level.replace('l', '');           // 'l2' → '2', 'l1' → '1'
            const btnId    = 'mondayL' + num + 'Btn';          // 'mondayL2Btn'
            const statusId = 'mondayStatusL' + num;            // 'mondayStatusL2'
            const btn      = document.getElementById(btnId);
            const statusEl = document.getElementById(statusId);

            if (!btn || !statusEl) {
                console.error('Monday button/status element not found:', btnId, statusId);
                return;
            }

            btn.disabled = true;
            btn.innerHTML = '<span class="spinner"></span>Uploading…';
            statusEl.className = 'monday-status pending';
            statusEl.textContent = '⏳ Upload in progress… this can take a bit for large files.';

            let uploadSuccess = false;
            try {
                const resp = await fetch('/monday_' + level, { method: 'POST' });
                const data = await resp.json();
                if (data.success) {
                    uploadSuccess = true;
                    statusEl.className = 'monday-status ok';
                    statusEl.textContent =
                        '✓ Monday L' + num + ' — created: ' + data.created +
                        '  skipped: ' + data.skipped +
                        (data.errors ? '  errors: ' + data.errors : '') +
                        '  (board ' + data.board + ')';
                } else {
                    statusEl.className = 'monday-status err';
                    statusEl.textContent = '✗ ' + (data.error || 'Upload failed');
                }
            } catch (err) {
                statusEl.className = 'monday-status err';
                statusEl.textContent = '✗ Network error: ' + err.message;
            }

            btn.disabled = false;
            const svgIcon = `<svg class="monday-logo" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg"><circle cx="7" cy="30" r="6" fill="#ff3d57"/><circle cx="20" cy="30" r="6" fill="#ffcc00"/><circle cx="33" cy="30" r="6" fill="#00ca72"/><path d="M4 24 L10 12" stroke="#ff3d57" stroke-width="3.5" stroke-linecap="round"/><path d="M17 24 L23 12" stroke="#ffcc00" stroke-width="3.5" stroke-linecap="round"/><path d="M30 24 L36 12" stroke="#00ca72" stroke-width="3.5" stroke-linecap="round"/></svg>`;
            btn.innerHTML = svgIcon + ' Upload to Monday L' + num;
            // Enable Sync L3 button after a successful L3 upload
            if (num === '3' && uploadSuccess) {
                const syncBtn = document.getElementById('mondaySyncL3Btn');
                if (syncBtn) syncBtn.disabled = false;
            }
        }

        async function syncL3() {
            const btn      = document.getElementById('mondaySyncL3Btn');
            const statusEl = document.getElementById('mondayStatusSyncL3');
            if (!btn || !statusEl) return;

            btn.disabled = true;
            btn.innerHTML = '<span class="spinner"></span>Syncing…';
            statusEl.className = 'monday-status pending';
            statusEl.textContent = '⏳ Sync in progress…';

            try {
                const resp = await fetch('/monday_l3_sync', { method: 'POST',
                    headers: {'Content-Type': 'application/json'}, body: '{}' });
                const data = await resp.json();
                if (data.success) {
                    statusEl.className = 'monday-status ok';
                    statusEl.textContent = '✓ Sync L3 — updated: ' + data.updated +
                        (data.errors ? '  errors: ' + data.errors : '') +
                        '  (group ' + data.group + ')';
                } else {
                    statusEl.className = 'monday-status err';
                    statusEl.textContent = '✗ ' + (data.error || 'Sync failed');
                }
            } catch (err) {
                statusEl.className = 'monday-status err';
                statusEl.textContent = '✗ Network error: ' + err.message;
            }

            btn.disabled = false;
            const svgIcon = `<svg class="monday-logo" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg"><circle cx="7" cy="30" r="6" fill="#ff3d57"/><circle cx="20" cy="30" r="6" fill="#ffcc00"/><circle cx="33" cy="30" r="6" fill="#00ca72"/><path d="M4 24 L10 12" stroke="#ff3d57" stroke-width="3.5" stroke-linecap="round"/><path d="M17 24 L23 12" stroke="#ffcc00" stroke-width="3.5" stroke-linecap="round"/><path d="M30 24 L36 12" stroke="#00ca72" stroke-width="3.5" stroke-linecap="round"/></svg>`;
            btn.innerHTML = svgIcon + ' Sync L3';
        }

        async function uploadEdited(input) {
            const statusEl = document.getElementById('editedUploadStatus');
            const file = input.files[0];
            if (!file) return;

            statusEl.className = 'monday-status';
            statusEl.className = 'monday-status pending';
            statusEl.textContent = '⏳ Uploading edited file…';

            const formData = new FormData();
            formData.append('file', file);

            try {
                const resp = await fetch('/upload_edited', { method: 'POST', body: formData });
                const data = await resp.json();
                if (data.success) {
                    statusEl.className = 'monday-status ok';
                    statusEl.textContent =
                        '✓ Edited file loaded — ' + data.rows + ' rows, ' +
                        data.columns + ' columns (' + data.type + '). ' +
                        'Monday uploads will now use this version.';
                } else {
                    statusEl.className = 'monday-status err';
                    statusEl.textContent = '✗ ' + (data.error || 'Upload failed');
                }
            } catch (err) {
                statusEl.className = 'monday-status err';
                statusEl.textContent = '✗ Network error: ' + err.message;
            }
            // Reset input so same file can be re-uploaded if needed
            input.value = '';
        }

        function showError(msg) {
            const card = document.getElementById('resultCard');
            card.className = 'result-card visible error-card';
            document.getElementById('resultIcon').className = 'result-icon err';
            document.getElementById('resultIcon').textContent = '✗';
            document.getElementById('resultTitle').textContent = 'Pipeline failed';
            document.getElementById('resultMeta').textContent = 'Check the log above for details';
            document.getElementById('resultContent').innerHTML = `<div class="error-text">${msg}</div>`;
        }

        // Allow Enter key
        document.getElementById('folderPath').addEventListener('keydown', e => {
            if (e.key === 'Enter') runPipeline();
        });

        /* ─── In-App Editor ─────────────────────────────────────────── */
        let editorColumns   = [];
        let lastCheckedRow  = null;   // for shift+click range selection
        let copiedRows      = [];     // clipboard for row copy/paste

        /* ── Open ─────────────────────────────────────────────────── */
        async function openEditor() {
            const overlay  = document.getElementById('editorOverlay');
            const statusEl = document.getElementById('editorStatus');
            const wrap     = document.getElementById('editorTableWrap');
            const meta     = document.getElementById('editorMeta');

            overlay.classList.add('open');
            document.body.style.overflow = 'hidden';
            statusEl.className = 'editor-statusbar-msg';
            statusEl.textContent = 'Loading data…';
            meta.textContent = '';
            wrap.innerHTML = '';
            clearColFilters();

            try {
                const resp = await fetch('/get_data');
                const data = await resp.json();
                if (!data.success) {
                    statusEl.className = 'editor-statusbar-msg err';
                    statusEl.textContent = '✗ ' + (data.error || 'Failed to load data');
                    return;
                }
                editorColumns = data.columns;
                renderEditorTable(data.columns, data.rows);
                meta.textContent = data.rows.length + ' rows · ' + data.columns.length + ' cols';
                statusEl.className = 'editor-statusbar-msg';
                statusEl.textContent = 'Ready — ' + data.rows.length + ' rows loaded';
                applyAllFilters();
            } catch (err) {
                statusEl.className = 'editor-statusbar-msg err';
                statusEl.textContent = '✗ Network error: ' + err.message;
            }
        }

        /* ── Render table ─────────────────────────────────────────── */
        function renderEditorTable(columns, rows) {
            const wrap  = document.getElementById('editorTableWrap');
            const table = document.createElement('table');
            table.className = 'editor-table';
            table.id = 'editorTable';

            // ── Header row 1: checkbox · del · column names
            const thead = table.createTHead();
            const hr1   = thead.insertRow();

            // Select-all checkbox
            const thChk = document.createElement('th');
            thChk.style.cssText = 'width:36px;text-align:center;';
            const chkAll = document.createElement('input');
            chkAll.type = 'checkbox';
            chkAll.title = 'Select / deselect all';
            chkAll.id = 'editorSelectAll';
            chkAll.onchange = () => toggleSelectAll(chkAll.checked);
            thChk.appendChild(chkAll);
            hr1.appendChild(thChk);

            // Del column header
            const thDel = document.createElement('th');
            thDel.style.cssText = 'width:36px;';
            hr1.appendChild(thDel);

            // Column name headers
            columns.forEach(col => {
                const th = document.createElement('th');
                th.textContent = col;
                th.title = col;
                hr1.appendChild(th);
            });

            // ── Header row 2: spacer · spacer · per-column filter inputs
            const hr2 = thead.insertRow();
            hr2.id = 'editorFilterRow';

            const fSpc1 = document.createElement('th');
            fSpc1.style.cssText = 'background:var(--layer-03);padding:4px;';
            hr2.appendChild(fSpc1);
            const fSpc2 = document.createElement('th');
            fSpc2.style.cssText = 'background:var(--layer-03);padding:4px;';
            hr2.appendChild(fSpc2);

            columns.forEach((col, ci) => {
                const th  = document.createElement('th');
                th.style.cssText = 'background:var(--layer-03);padding:3px 4px;';
                const inp = document.createElement('input');
                inp.type = 'text';
                inp.placeholder = '▾ filter';
                inp.dataset.colIdx = ci;
                inp.className = 'col-filter-input';
                inp.style.cssText = 'width:100%;min-width:60px;padding:3px 6px;font-size:11px;' +
                    'background:var(--layer-01);border:1px solid var(--border-subtle);border-radius:2px;' +
                    'color:var(--text-primary);font-family:inherit;outline:none;box-sizing:border-box;';
                inp.oninput = () => applyAllFilters();
                th.appendChild(inp);
                hr2.appendChild(th);
            });

            // ── Body
            const tbody = table.createTBody();
            rows.forEach((row) => {
                appendEditorRow(tbody, columns, row);
            });

            wrap.innerHTML = '';
            wrap.appendChild(table);
            lastCheckedRow = null;
        }

        function appendEditorRow(tbody, columns, rowData) {
            const tr = tbody.insertRow();

            // Checkbox cell
            const tdChk = tr.insertCell();
            tdChk.style.cssText = 'width:36px;text-align:center;padding:4px;background:var(--layer-02);';
            const chk = document.createElement('input');
            chk.type = 'checkbox';
            chk.className = 'row-check';
            chk.onclick = (e) => handleRowCheck(e, tr);
            tdChk.appendChild(chk);

            // Delete cell
            const tdDel = tr.insertCell();
            tdDel.className = 'del-cell';
            const delBtn = document.createElement('button');
            delBtn.className = 'editor-del-btn';
            delBtn.textContent = '✕';
            delBtn.title = 'Delete row';
            delBtn.onclick = () => { tr.remove(); applyAllFilters(); updateSelectionUI(); };
            tdDel.appendChild(delBtn);

            // Data cells
            (rowData || []).forEach((val, ci) => {
                const td = tr.insertCell();
                td.contentEditable = 'true';
                td.spellcheck = false;
                td.dataset.col = ci;
                td.textContent = (val === null || val === undefined) ? '' : String(val);
                // update search cache on edit
                td.oninput = () => refreshRowSearchCache(tr);
            });

            refreshRowSearchCache(tr);
            return tr;
        }

        function refreshRowSearchCache(tr) {
            const cells = Array.from(tr.cells).slice(2); // skip checkbox + del
            tr.dataset.searchText = cells.map(td => td.textContent).join('|').toLowerCase();
        }

        /* ── Selection ────────────────────────────────────────────── */
        function handleRowCheck(e, tr) {
            if (e.shiftKey && lastCheckedRow) {
                const tbody = tr.closest('tbody');
                const rows  = Array.from(tbody.rows);
                const a = rows.indexOf(lastCheckedRow);
                const b = rows.indexOf(tr);
                const [lo, hi] = a < b ? [a, b] : [b, a];
                const state = tr.querySelector('.row-check').checked;
                rows.slice(lo, hi + 1).forEach(r => {
                    const cb = r.querySelector('.row-check');
                    if (cb) cb.checked = state;
                });
            }
            lastCheckedRow = tr;
            updateSelectionUI();
        }

        function toggleSelectAll(checked) {
            const tbody = document.querySelector('#editorTable tbody');
            if (!tbody) return;
            Array.from(tbody.rows).forEach(tr => {
                if (tr.classList.contains('hidden-row')) return;
                const cb = tr.querySelector('.row-check');
                if (cb) cb.checked = checked;
            });
            updateSelectionUI();
        }

        function getSelectedRows() {
            return Array.from(document.querySelectorAll('#editorTable tbody tr:not(.hidden-row) .row-check:checked'))
                .map(cb => cb.closest('tr'));
        }

        function updateSelectionUI() {
            const n   = getSelectedRows().length;
            const btn = document.getElementById('editorDeleteSelBtn');
            const cpBtn = document.getElementById('editorCopySelBtn');
            const lbl = document.getElementById('editorSelCount');
            if (btn)  { btn.disabled = n === 0; btn.textContent  = n > 0 ? 'Delete Selected (' + n + ')' : 'Delete Selected'; }
            if (cpBtn){ cpBtn.disabled = n === 0; }
            if (lbl)  { lbl.textContent = n > 0 ? n + ' selected' : ''; }
            // sync select-all checkbox state
            const tbody = document.querySelector('#editorTable tbody');
            if (tbody) {
                const visRows = Array.from(tbody.rows).filter(r => !r.classList.contains('hidden-row'));
                const allChk = document.getElementById('editorSelectAll');
                if (allChk) allChk.checked = visRows.length > 0 && visRows.every(r => r.querySelector('.row-check')?.checked);
            }
        }

        function deleteSelected() {
            getSelectedRows().forEach(tr => tr.remove());
            applyAllFilters();
            updateSelectionUI();
        }

        /* ── Copy / Paste rows ────────────────────────────────────── */
        function copySelectedRows() {
            copiedRows = getSelectedRows().map(tr =>
                Array.from(tr.cells).slice(2).map(td => td.textContent)
            );
            setStatus('Copied ' + copiedRows.length + ' row(s) — use Paste Rows to insert below last selected row.');
        }

        function pasteRows() {
            if (!copiedRows.length) return;
            const tbody   = document.querySelector('#editorTable tbody');
            if (!tbody) return;
            const selRows = getSelectedRows();
            const anchor  = selRows.length ? selRows[selRows.length - 1] : null;

            copiedRows.forEach(rowData => {
                const newTr = appendEditorRow(tbody, editorColumns, rowData);
                if (anchor) anchor.after(newTr);
                else tbody.appendChild(newTr);
            });
            applyAllFilters();
            updateSelectionUI();
            setStatus('Pasted ' + copiedRows.length + ' row(s).');
        }

        /* ── Filter ───────────────────────────────────────────────── */
        function clearColFilters() {
            document.querySelectorAll('.col-filter-input').forEach(inp => inp.value = '');
        }

        function applyAllFilters() {
            const globalQ   = (document.getElementById('editorGlobalSearch')?.value || '').trim().toLowerCase();
            const colInputs = Array.from(document.querySelectorAll('.col-filter-input'));
            const colFilters = colInputs.map(inp => inp.value.trim().toLowerCase());
            const hasColFilter = colFilters.some(f => f !== '');

            const tbody = document.querySelector('#editorTable tbody');
            if (!tbody) return;

            Array.from(tbody.rows).forEach(tr => {
                let show = true;
                const text = tr.dataset.searchText || '';

                if (globalQ && !text.includes(globalQ)) show = false;

                if (show && hasColFilter) {
                    const cells = Array.from(tr.cells).slice(2);
                    colFilters.forEach((f, i) => {
                        if (f && cells[i] && !cells[i].textContent.toLowerCase().includes(f)) show = false;
                    });
                }

                tr.classList.toggle('hidden-row', !show);
            });

            updateRowCount();
            updateSelectionUI();
        }

        function clearAllFilters() {
            clearColFilters();
            const g = document.getElementById('editorGlobalSearch');
            if (g) g.value = '';
            applyAllFilters();
        }

        /* ── Row count ─────────────────────────────────────────────── */
        function updateRowCount() {
            const tbody = document.querySelector('#editorTable tbody');
            if (!tbody) return;
            const total   = tbody.rows.length;
            const visible = Array.from(tbody.rows).filter(r => !r.classList.contains('hidden-row')).length;
            const el = document.getElementById('editorRowCount');
            if (el) el.textContent = visible < total ? visible + ' of ' + total + ' rows' : total + ' rows';
        }

        /* ── Keyboard: Tab/Enter navigation, Ctrl+D duplicate ─────── */
        document.addEventListener('keydown', e => {
            const overlay = document.getElementById('editorOverlay');
            if (!overlay.classList.contains('open')) return;

            // Esc: close
            if (e.key === 'Escape') { closeEditor(); return; }

            // Ctrl+S: save
            if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); saveEditorData(); return; }

            const active = document.activeElement;
            if (!active || !active.matches('#editorTable td[contenteditable]')) return;

            const tr      = active.closest('tr');
            const tbody   = tr?.closest('tbody');
            const allCells = tbody
                ? Array.from(tbody.rows)
                      .filter(r => !r.classList.contains('hidden-row'))
                      .flatMap(r => Array.from(r.cells).slice(2).filter(td => td.contentEditable === 'true'))
                : [];
            const idx = allCells.indexOf(active);
            const colCount = editorColumns.length;

            // Tab: next cell  /  Shift+Tab: prev cell
            if (e.key === 'Tab') {
                e.preventDefault();
                const next = allCells[idx + (e.shiftKey ? -1 : 1)];
                if (next) { next.focus(); placeCaretAtEnd(next); }
                return;
            }

            // Enter: cell below
            if (e.key === 'Enter' && !e.shiftKey) {
                e.preventDefault();
                const next = allCells[idx + colCount];
                if (next) { next.focus(); placeCaretAtEnd(next); }
                return;
            }

            // Arrow keys: move between cells
            if (e.key === 'ArrowRight' && isCaretAtEnd(active)) {
                e.preventDefault();
                const next = allCells[idx + 1]; if (next) { next.focus(); placeCaretAtStart(next); }
                return;
            }
            if (e.key === 'ArrowLeft' && isCaretAtStart(active)) {
                e.preventDefault();
                const prev = allCells[idx - 1]; if (prev) { prev.focus(); placeCaretAtEnd(prev); }
                return;
            }
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                const next = allCells[idx + colCount]; if (next) { next.focus(); placeCaretAtStart(next); }
                return;
            }
            if (e.key === 'ArrowUp') {
                e.preventDefault();
                const prev = allCells[idx - colCount]; if (prev) { prev.focus(); placeCaretAtStart(prev); }
                return;
            }

            // Ctrl+D: duplicate row
            if ((e.ctrlKey || e.metaKey) && e.key === 'd') {
                e.preventDefault();
                const rowData = Array.from(tr.cells).slice(2).map(td => td.textContent);
                const newTr   = appendEditorRow(tbody, editorColumns, rowData);
                tr.after(newTr);
                applyAllFilters();
                // focus first data cell of new row
                const firstCell = newTr.cells[2];
                if (firstCell) firstCell.focus();
            }
        });

        function placeCaretAtEnd(el) {
            const range = document.createRange();
            const sel   = window.getSelection();
            range.selectNodeContents(el);
            range.collapse(false);
            sel.removeAllRanges(); sel.addRange(range);
        }
        function placeCaretAtStart(el) {
            const range = document.createRange();
            const sel   = window.getSelection();
            range.selectNodeContents(el);
            range.collapse(true);
            sel.removeAllRanges(); sel.addRange(range);
        }
        function isCaretAtEnd(el) {
            const sel = window.getSelection();
            if (!sel.rangeCount) return true;
            const range = sel.getRangeAt(0);
            return range.collapsed && range.endOffset === (el.textContent || '').length;
        }
        function isCaretAtStart(el) {
            const sel = window.getSelection();
            if (!sel.rangeCount) return true;
            return sel.getRangeAt(0).startOffset === 0;
        }

        /* ── Close ─────────────────────────────────────────────────── */
        function closeEditor() {
            document.getElementById('editorOverlay').classList.remove('open');
            document.body.style.overflow = '';
        }

        /* ── Status helper ─────────────────────────────────────────── */
        function setStatus(msg, cls) {
            const el = document.getElementById('editorStatus');
            if (!el) return;
            el.className = 'editor-statusbar-msg' + (cls ? ' ' + cls : '');
            el.textContent = msg;
        }

        /* ── Save ──────────────────────────────────────────────────── */
        async function saveEditorData() {
            const tbody = document.querySelector('#editorTable tbody');
            if (!tbody) return;
            setStatus('⏳ Saving…');

            const allRows = [];
            Array.from(tbody.rows).forEach(tr => {
                const cells = Array.from(tr.cells).slice(2);
                allRows.push(cells.map(td => td.textContent));
            });

            try {
                const resp = await fetch('/save_data', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({ columns: editorColumns, rows: allRows })
                });
                const data = await resp.json();
                if (data.success) {
                    setStatus('✓ Saved — ' + data.rows + ' rows · ' + data.columns + ' columns. Monday uploads and Download now use this version.', 'ok');
                    document.getElementById('editorMeta').textContent = data.rows + ' rows · ' + data.columns + ' cols';
                } else {
                    setStatus('✗ ' + (data.error || 'Save failed'), 'err');
                }
            } catch (err) {
                setStatus('✗ Network error: ' + err.message, 'err');
            }
        }
    </script>

    <!-- In-App Editor (full-screen) -->
    <div id="editorOverlay" class="editor-overlay">

        <!-- ── Top bar ───────────────────────────────────────────── -->
        <div class="editor-topbar">
            <span class="editor-topbar-title">Inspect &amp; Edit</span>
            <span class="editor-topbar-meta" id="editorMeta"></span>
            <button class="editor-save-btn" onclick="saveEditorData()">&#10003;&nbsp; Save Changes</button>
            <button class="editor-close-btn" onclick="closeEditor()">&#10005;&nbsp; Close</button>
        </div>

        <!-- ── Action / filter toolbar ──────────────────────────── -->
        <div class="editor-toolbar2">
            <!-- Global search -->
            <input id="editorGlobalSearch" type="text" class="editor-global-search"
                   placeholder="&#128269;  Search all columns…" oninput="applyAllFilters()">

            <!-- Divider -->
            <div class="editor-tb-divider"></div>

            <!-- Batch actions -->
            <button id="editorDeleteSelBtn" class="editor-tb-btn danger" onclick="deleteSelected()" disabled>
                &#128465;&nbsp; Delete Selected
            </button>
            <button id="editorCopySelBtn" class="editor-tb-btn" onclick="copySelectedRows()" disabled>
                &#128203;&nbsp; Copy Rows
            </button>
            <button class="editor-tb-btn" onclick="pasteRows()" title="Paste copied rows below last selected row">
                &#128204;&nbsp; Paste Rows
            </button>

            <!-- Divider -->
            <div class="editor-tb-divider"></div>

            <!-- Clear filters -->
            <button class="editor-tb-btn" onclick="clearAllFilters()">
                &#10005;&nbsp; Clear Filters
            </button>

            <!-- Row / selection count -->
            <span class="editor-row-count" id="editorRowCount"></span>
            <span class="editor-sel-count" id="editorSelCount"></span>
        </div>

        <!-- ── Scrollable table ──────────────────────────────────── -->
        <div id="editorTableWrap" class="editor-table-wrap"></div>

        <!-- ── Status bar ───────────────────────────────────────── -->
        <div class="editor-statusbar">
            <span class="editor-statusbar-msg" id="editorStatus"></span>
            <span class="editor-statusbar-hint">
                Click cell to edit &nbsp;·&nbsp;
                Tab / Enter / Arrows to navigate &nbsp;·&nbsp;
                Ctrl+D duplicate row &nbsp;·&nbsp;
                Shift+click to select range &nbsp;·&nbsp;
                Ctrl+S to save
            </span>
        </div>
    </div>
</body>
</html>
"""

# In-memory store for the combined file and processed DataFrame
_output_buffer  = None
_output_meta    = {}
_combined_df    = None   # processed DataFrame — used by Monday upload routes
_combined_type  = ""     # "p2f" | "qfs" | "auto"
_last_l3_group  = ""     # group_id of the most recently uploaded L3 group

def _normalize_col_name(c: object) -> str:
    s = str(c) if c is not None else ""
    s = s.strip().lower()
    # Replace any non-word char (incl spaces) with underscore.
    s = re.sub(r"[^\w]+", "_", s, flags=re.UNICODE)
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "col"


def _dedupe_columns(cols) -> list[str]:
    """Ensure unique column names by suffixing _2, _3, ...

    Handles the edge case where a pre-existing column already has the
    same name as a generated suffix (e.g. col_2 already exists when
    a second 'col' would otherwise be renamed to col_2).
    """
    used: set[str] = set()
    seen: dict[str, int] = {}
    out: list[str] = []
    for c in cols:
        base = str(c)
        if base not in used:
            used.add(base)
            seen[base] = 1
            out.append(base)
        else:
            n = seen.get(base, 0) + 1
            candidate = f"{base}_{n}"
            while candidate in used:
                n += 1
                candidate = f"{base}_{n}"
            seen[base] = n
            used.add(candidate)
            out.append(candidate)
    return out


def _strip_object_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Strip whitespace for object/string-like columns.

    Important: if the dataframe has duplicate column names, `df[col]` returns a DataFrame
    (not a Series) which breaks `.str`. So we operate positionally.
    """
    if df.empty:
        return df

    for i, dtype in enumerate(df.dtypes):
        if dtype == object or str(dtype).startswith("string"):
            s = df.iloc[:, i].astype("string").str.strip()
            df.iloc[:, i] = s
    return df


def _drop_effectively_empty_rows(df: pd.DataFrame) -> pd.DataFrame:
    """
    Drop rows that are empty/whitespace across all columns.
    (Much faster than row-wise apply on large frames.)
    """
    if df.empty:
        return df
    tmp = df.copy()
    obj_pos = [i for i, dtype in enumerate(tmp.dtypes) if dtype == object or str(dtype).startswith("string")]
    if len(obj_pos) > 0:
        tmp.iloc[:, obj_pos] = tmp.iloc[:, obj_pos].replace(r"^\s*$", pd.NA, regex=True)
    return df.loc[~tmp.isna().all(axis=1)]


def run_processor(dfs, logs):
    """Step 1: Basic processing — strip whitespace, normalize columns."""
    logs.append({"msg": "processor · normalizing column names & whitespace", "type": "info", "step": "processor", "step_state": "running"})
    processed = []
    for df in dfs:
        cols = [_normalize_col_name(c) for c in df.columns]
        df.columns = _dedupe_columns(cols)
        df = df.dropna(how="all")
        df = _strip_object_columns(df)
        processed.append(df)
    logs.append({"msg": f"processor · processed {len(processed)} dataframes", "type": "", "step": "processor", "step_state": "done"})
    return processed


def run_quality(dfs, logs):
    """Step 2: Data quality — flag/remove empty rows, deduplicate."""
    logs.append({"msg": "quality · running quality checks", "type": "info", "step": "quality", "step_state": "running"})
    cleaned = []
    total_removed = 0
    for df in dfs:
        before = len(df)
        df = _drop_effectively_empty_rows(df)
        df = df.drop_duplicates()
        after = len(df)
        removed = before - after
        total_removed += removed
        cleaned.append(df)
    logs.append({"msg": f"quality · removed {total_removed} duplicate/empty rows", "type": "warn" if total_removed else "", "step": "quality", "step_state": "done"})
    return cleaned


def run_checker(dfs, logs):
    """Step 3: Check schema consistency across files."""
    logs.append({"msg": "checker · validating schema consistency", "type": "info", "step": "checker", "step_state": "running"})
    if not dfs:
        logs.append({"msg": "checker · no dataframes to check", "type": "warn", "step": "checker", "step_state": "done"})
        return dfs

    base_cols = set(dfs[0].columns)
    mismatches = 0
    for i, df in enumerate(dfs[1:], 2):
        extra = set(df.columns) - base_cols
        missing = base_cols - set(df.columns)
        if extra:
            logs.append({"msg": f"checker · file {i}: extra cols {list(extra)}", "type": "warn"})
        if missing:
            logs.append({"msg": f"checker · file {i}: missing cols {list(missing)}", "type": "warn"})
            mismatches += 1

    if mismatches == 0:
        logs.append({"msg": "checker · all schemas consistent ✓", "type": "success", "step": "checker", "step_state": "done"})
    else:
        logs.append({"msg": f"checker · {mismatches} schema mismatch(es) — will align on combine", "type": "warn", "step": "checker", "step_state": "done"})
    return dfs


def run_combiner(dfs, logs):
    """Step 4: Combine all dataframes."""
    logs.append({"msg": "combiner · merging all dataframes", "type": "info", "step": "combiner", "step_state": "running"})
    if not dfs:
        return pd.DataFrame()

    # Align schemas on the union of columns (deterministic order).
    all_cols = []
    seen = set()
    for df in dfs:
        for c in df.columns:
            if c not in seen:
                seen.add(c)
                all_cols.append(c)

    # Prefer meta columns first if present.
    meta_first = [c for c in ["_source_file", "_source_sheet"] if c in seen]
    rest = [c for c in all_cols if c not in set(meta_first)]
    ordered_cols = meta_first + rest

    aligned = [df.reindex(columns=ordered_cols) for df in dfs]
    combined = pd.concat(aligned, ignore_index=True, sort=False, copy=False)
    logs.append({"msg": f"combiner · combined {len(dfs)} files → {len(combined)} rows, {len(combined.columns)} cols", "type": "success", "step": "combiner", "step_state": "done"})
    return combined


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/run_pipeline', methods=['POST'])
def run_pipeline():
    global _output_buffer, _output_meta

    data = request.get_json()
    folder_path = (data or {}).get('folder_path', '').strip()
    logs = []
    start = datetime.now()

    if not os.path.isdir(folder_path):
        return jsonify({"success": False, "error": f"Folder not found:\n{folder_path}", "logs": logs, "failed_step": "processor"})

    try:
        # Step 1–4 (via src.combiner + existing cleanup)
        logs.append({"msg": "combiner · loading + consolidating via src/combiner.py", "type": "info", "step": "combiner", "step_state": "running"})
        result = combine_folder_to_frames(folder_path)
        combined = result["combined_df"]
        summary = result["summary_df"]
        meta = result["meta"]
        logs.append({"msg": f"combiner · combined {meta['files_loaded']} file(s) → {meta['total_rows']} rows, {meta['total_cols']} cols", "type": "success", "step": "combiner", "step_state": "done"})

        # Keep existing processor/quality passes for normalization + de-dupe
        logs.append({"msg": "processor · normalizing column names & whitespace", "type": "info", "step": "processor", "step_state": "running"})
        combined.columns = _dedupe_columns([_normalize_col_name(c) for c in combined.columns])
        combined = combined.dropna(how="all")
        combined = _strip_object_columns(combined)
        logs.append({"msg": "processor · done", "type": "", "step": "processor", "step_state": "done"})

        logs.append({"msg": "quality · dropping empty rows (no de-duplication)", "type": "info", "step": "quality", "step_state": "running"})
        before = len(combined)
        combined = _drop_effectively_empty_rows(combined)
        removed = before - len(combined)
        logs.append({"msg": f"quality · removed {removed} empty rows", "type": "warn" if removed else "", "step": "quality", "step_state": "done"})

        logs.append({"msg": "checker · schema consistency skipped (union schema used)", "type": "dim", "step": "checker", "step_state": "done"})

        # Detect P2F schema: either the "Number" column variant (FSD/PDD files) or
        # the plain "Criteria ID – Deliverable Name" variant (CRD files in P2F MFG).
        # CRD P2F files are distinguished from Q&FS by having "CRD" in the source filenames.
        _has_p2f_number_col = any(c in combined.columns for c in [
            "criteria_id_number_deliverable_name", "Criteria ID Number \u2013 Deliverable Name"
        ])
        _has_crd_col = "criteria_id_deliverable_name" in combined.columns
        _src_col = next((c for c in ["source_file", "_source_file"] if c in combined.columns), None)
        _src_is_crd = (
            _has_crd_col and _src_col is not None
            and combined[_src_col].astype("string").str.contains(
                r"(?i)CRD_", regex=True, na=False
            ).any()
        )
        _is_p2f = _has_p2f_number_col or _src_is_crd

        # If this looks like the P2F dataset, reshape output to match the reference P2F_Final.xlsx
        if _is_p2f:
            logs.append({"msg": "formatter · shaping output to P2F_Final schema", "type": "info"})
            combined = to_p2f_final_schema(combined)

        # If this looks like the Q&FS dataset (review_date + criteria columns, no P2F marker),
        # reshape output to match the reference Q&FS_cleaned_final schema.
        elif (
            "review_date" in combined.columns
            and any(c in combined.columns for c in [
                "criteria_id_deliverable_name", "criteria_id", "criteria_name", "nt", "1"
            ])
        ):
            logs.append({"msg": "formatter · shaping output to Q&FS schema", "type": "info"})
            combined = to_qfs_final_schema(combined)

        # Step 5: Export
        logs.append({"msg": "export · writing to Excel buffer", "type": "info", "step": "export", "step_state": "running"})
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            # P2F schema: "Criteria ID Number – Deliverable Name" column present
            is_p2f = "Criteria ID Number \u2013 Deliverable Name" in combined.columns
            # Q&FS schema: has "Criteria ID" (short form) + "Deliverable VersionName"
            is_qfs = (
                not is_p2f
                and "Criteria ID" in combined.columns
                and "Deliverable VersionName" in combined.columns
            )
            use_full_data = is_p2f or is_qfs
            main_sheet = "Full Data" if use_full_data else "Combined"
            combined.to_excel(writer, index=False, sheet_name=main_sheet)

            # Both P2F and Q&FS get a Checks sheet (score aggregated per deliverable)
            if use_full_data and "Deliverable VersionName" in combined.columns and "Score" in combined.columns:
                checks_rows = []
                for i, (dvn, grp) in enumerate(
                    combined.groupby("Deliverable VersionName", sort=False), start=1
                ):
                    total_score = grp["Score"].apply(pd.to_numeric, errors="coerce").sum()
                    checks_rows.append({
                        "#": i,
                        "Deliverable / File Name": dvn,
                        "Original Score": total_score,
                        "Combined File Score": total_score,
                    })
                checks_df = pd.DataFrame(checks_rows)
                checks_df.to_excel(writer, index=False, sheet_name="Checks")

            # Lightweight summary sheet for traceability (non-schema outputs only).
            if summary is not None and not use_full_data:
                summary.to_excel(writer, index=False, sheet_name="Summary")
        buf.seek(0)

        _output_buffer = buf.read()
        duration = round((datetime.now() - start).total_seconds(), 2)
        logs.append({"msg": f"export · done in {duration}s", "type": "success"})

        _output_meta = {
            "files_processed": meta.get("files_loaded"),
            "total_rows": len(combined),
            "total_cols": len(combined.columns),
            "duration": duration
        }

        # Store processed DataFrame for on-demand Monday uploads via UI buttons
        global _combined_df, _combined_type
        _monday_type   = "p2f" if _is_p2f else ("qfs" if "review_date" in combined.columns else "auto")
        _combined_df   = combined.copy()
        _combined_type = _monday_type

        return jsonify({
            "success": True,
            "logs": logs,
            **_output_meta
        })

    except Exception as e:
        tb = traceback.format_exc()
        logs.append({"msg": "ERROR: " + str(e), "type": "error"})
        return jsonify({"success": False, "error": str(e) + "\n\n" + tb, "logs": logs, "failed_step": "combiner"})


def _load_monday_module():
    """Dynamically load upload_to_monday so app.py has no hard import dependency on it."""
    import importlib.util, sys as _sys
    script = os.path.join(_ROOT, "upload_to_monday.py")
    script = os.path.normpath(script)
    spec = importlib.util.spec_from_file_location("upload_to_monday", script)
    mod  = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@app.route('/monday_l2', methods=['POST'])
def monday_l2():
    """Upload unique Deliverable VersionNames to Monday.com L2 board."""
    global _combined_df, _combined_type
    if _combined_df is None:
        return jsonify({"success": False, "error": "Run the pipeline first before uploading."})
    try:
        mod = _load_monday_module()
        cfg = mod.load_config()
        stats = mod.upload_unique_names(_combined_df, _combined_type, cfg, dry_run=False)
        return jsonify({
            "success": True,
            "created": stats["created"],
            "skipped": stats["skipped"],
            "errors":  stats["errors"],
            "board":   cfg.get("board_id", ""),
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route('/monday_l1', methods=['POST'])
def monday_l1():
    """Upload unique Deliverable VersionNames to Monday.com L1 board."""
    global _combined_df, _combined_type
    if _combined_df is None:
        return jsonify({"success": False, "error": "Run the pipeline first before uploading."})
    try:
        mod = _load_monday_module()
        cfg = mod.load_config()
        l1_board = cfg.get("l1_board_id", "")
        if not l1_board or l1_board == "L1_BOARD_ID_HERE":
            return jsonify({"success": False,
                            "error": "L1 board ID not set. Add 'l1_board_id' to monday_config.json."})
        # Temporarily swap board_id to the L1 board for this upload
        cfg_l1 = dict(cfg)
        cfg_l1["board_id"]  = l1_board
        cfg_l1["group_map"] = {k: v for k, v in cfg.get("l1_group_map", cfg["group_map"]).items()
                                if not k.startswith("_")}
        stats = mod.upload_unique_names(_combined_df, _combined_type, cfg_l1, dry_run=False)
        return jsonify({
            "success": True,
            "created": stats["created"],
            "skipped": stats["skipped"],
            "errors":  stats["errors"],
            "board":   l1_board,
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route('/monday_l3', methods=['POST'])
def monday_l3():
    """Upload every criteria row to Monday.com L3 board in a new dated group."""
    global _combined_df, _combined_type, _last_l3_group
    if _combined_df is None:
        return jsonify({"success": False, "error": "Run the pipeline first before uploading."})
    try:
        mod = _load_monday_module()
        cfg = mod.load_config()
        l3_board = cfg.get("l3_board_id", "")
        if not l3_board or l3_board == "L3_BOARD_ID_HERE":
            return jsonify({"success": False,
                            "error": "L3 board ID not set. Add 'l3_board_id' to monday_config.json."})
        stats = mod.upload_to_l3_board(_combined_df, _combined_type, cfg, dry_run=False)
        _last_l3_group = stats.get("group_id", "")
        return jsonify({
            "success":  True,
            "created":  stats["created"],
            "skipped":  stats["skipped"],
            "errors":   stats["errors"],
            "board":    l3_board,
            "group_id": _last_l3_group,
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route('/monday_l3_sync', methods=['POST'])
def monday_l3_sync():
    """Trigger the 'Synch Version and criteria' status on all items in the last L3 group."""
    global _last_l3_group
    if not _last_l3_group:
        return jsonify({"success": False,
                        "error": "No L3 group found — run Upload to Monday L3 first."})
    try:
        mod = _load_monday_module()
        cfg = mod.load_config()
        l3_board  = cfg.get("l3_board_id", "")
        l3_sync   = cfg.get("l3_sync", {})
        col_id    = l3_sync.get("column_id", "color_mkwbaxzj")
        label     = (request.json.get("label") if request.is_json else None) \
                    or l3_sync.get("sync_label") or None
        stats = mod.sync_l3_group(cfg["api_token"], l3_board, _last_l3_group,
                                   column_id=col_id, label=label, dry_run=False)
        return jsonify({
            "success": True,
            "updated": stats["updated"],
            "errors":  stats["errors"],
            "group":   _last_l3_group,
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route('/download')
def download():
    global _output_buffer
    if not _output_buffer:
        return "No output available. Run the pipeline first.", 400
    buf = io.BytesIO(_output_buffer)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name='combined_output.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/get_data')
def get_data():
    """Return the current in-memory DataFrame as JSON for the in-app editor."""
    global _combined_df
    if _combined_df is None:
        return jsonify({"success": False, "error": "No data — run the pipeline first."})
    df = _combined_df.copy().fillna("")
    return jsonify({
        "success": True,
        "columns": list(df.columns),
        "rows":    df.values.tolist(),
    })


@app.route('/save_data', methods=['POST'])
def save_data():
    """Accept edited table data from the in-app editor and update _combined_df."""
    global _combined_df, _combined_type, _output_buffer
    if _combined_df is None:
        return jsonify({"success": False, "error": "No pipeline data to update."})
    try:
        body    = request.get_json(force=True)
        columns = body.get("columns", [])
        rows    = body.get("rows", [])
        if not columns or not rows:
            return jsonify({"success": False, "error": "Empty data received."})

        import numpy as np
        df = pd.DataFrame(rows, columns=columns)
        # Restore numeric types for Score column
        for col in df.columns:
            if col.lower() == "score":
                df[col] = pd.to_numeric(df[col], errors="coerce")
        # Replace empty strings with NaN
        df.replace("", pd.NA, inplace=True)

        _combined_df = df
        if not _combined_type:
            _combined_type = (
                "p2f" if any("Criteria ID Number" in str(c) for c in df.columns) else "qfs"
            )

        # Regenerate output buffer so Download reflects edits
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)
        _output_buffer = buf.getvalue()

        return jsonify({"success": True, "rows": len(df), "columns": len(df.columns)})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


@app.route('/upload_edited', methods=['POST'])
def upload_edited():
    """Accept an edited Excel file and replace the in-memory DataFrame for Monday uploads."""
    global _combined_df, _combined_type, _output_buffer
    if 'file' not in request.files:
        return jsonify({"success": False, "error": "No file provided."})
    f = request.files['file']
    if not f.filename:
        return jsonify({"success": False, "error": "Empty filename."})
    if not f.filename.lower().endswith(('.xlsx', '.xls')):
        return jsonify({"success": False, "error": "Only .xlsx / .xls files are accepted."})
    try:
        raw = f.read()
        df  = pd.read_excel(io.BytesIO(raw))
        if df.empty:
            return jsonify({"success": False, "error": "The uploaded file has no data rows."})

        _combined_df   = df
        # Keep the same dataset type; re-detect if needed
        if not _combined_type:
            _combined_type = (
                "p2f" if any("Criteria ID Number" in str(c) for c in df.columns) else "qfs"
            )
        # Also replace output buffer so Download reflects the edited file
        _output_buffer = raw

        cols    = list(df.columns)
        preview = df.head(3).fillna("").to_dict(orient="records")
        return jsonify({
            "success":  True,
            "rows":     len(df),
            "columns":  len(cols),
            "col_names": cols[:10],   # first 10 column names for confirmation
            "preview":  preview,
            "type":     _combined_type.upper(),
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)})


if __name__ == '__main__':
    print("\n  IBM Excel Pipeline Server")
    print("  → http://localhost:5000\n")
    app.run(debug=True, port=5000)