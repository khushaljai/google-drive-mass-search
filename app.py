#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Drive File Search & Download (Streamlit)
- Upload an Excel with columns: filename, company
- Search Google Drive for each filename
- Generate per-company "Found / Not Found" results workbook
- Optionally download matched files into company subfolders

Configuration:
- Reads CONFIG_PATH env var or ./config.json
- Sample config:
{
  "credentials_file": "credentials.json",
  "token_file": "token.pickle",
  "scopes": ["https://www.googleapis.com/auth/drive.readonly"],
  "exclusion_suffixes": ["_backup", "_copy", "_old"],
  "download_folder": "./downloads"
}
"""

import os
import io
import re
import json
import time
import pickle
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Tuple

import pandas as pd
import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ---------------------------
# Config Helpers
# ---------------------------

def load_config() -> Dict:
    """Load configuration from CONFIG_PATH env var or local config.json."""
    cfg_path = os.getenv("CONFIG_PATH", "config.json")
    cfg_path = Path(cfg_path)
    if not cfg_path.exists():
        st.stop()
    with cfg_path.open("r", encoding="utf-8") as f:
        cfg = json.load(f)
    # sensible fallbacks
    cfg.setdefault("credentials_file", "credentials.json")
    cfg.setdefault("token_file", "token.pickle")
    cfg.setdefault("scopes", ["https://www.googleapis.com/auth/drive.readonly"])
    cfg.setdefault("exclusion_suffixes", ["_backup", "_copy", "_old"])
    cfg.setdefault("download_folder", "./downloads")
    return cfg

# ---------------------------
# Google Auth
# ---------------------------

def get_credentials(token_file: str, credentials_file: str, scopes: List[str]):
    """Return Google Credentials object using local token file & credentials.json."""
    creds = None
    token_path = Path(token_file)
    if token_path.exists():
        try:
            with token_path.open("rb") as token:
                creds = pickle.load(token)
        except Exception:
            creds = None

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception:
            creds = None

    if not creds:
        flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
        creds = flow.run_local_server(port=0)
        with token_path.open("wb") as token:
            pickle.dump(creds, token)

    return creds

def build_drive_service(creds):
    """Build and return Google Drive API service."""
    return build("drive", "v3", credentials=creds)

# ---------------------------
# Search & Download Helpers
# ---------------------------

def normalize_filename(name: str) -> str:
    """
    Normalize filename to improve matching:
      - remove bracket counters like " (1)" or " (2)"
      - strip whitespace
      - lower case
    """
    if not isinstance(name, str):
        return ""
    # remove trailing " (n)" counters
    name = re.sub(r"\s*\(\d+\)\s*$", "", name.strip())
    return name.strip().lower()

def should_exclude(name_no_ext: str, exclusion_suffixes: List[str]) -> bool:
    """
    Return True if the base name ends with any configured excluded suffix.
    Example: "report_copy" endswith "_copy" -> exclude
    """
    return any(name_no_ext.endswith(suf.lower()) for suf in exclusion_suffixes)

def drive_search_by_name(service, filename: str, max_results: int = 20) -> List[Dict]:
    """
    Search Drive files by 'name contains' strategy.
    Returns list of file dicts with id, name, mimeType.
    """
    # Escape quotes in filename for query safety
    safe = filename.replace('"', '\\"')
    query = f'name contains "{safe}" and trashed = false'
    results = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name, mimeType, size, webViewLink)",
        pageSize=max_results
    ).execute()
    return results.get("files", [])

def best_match_from_candidates(filename: str, candidates: List[Dict], exclusion_suffixes: List[str]) -> Dict:
    """
    From candidate list, pick the best match:
      - exact name match (case-insensitive) first
      - otherwise, the first candidate that isn't in excluded suffixes
      - otherwise, first candidate
    """
    target_norm = normalize_filename(Path(filename).stem)
    ext = Path(filename).suffix.lower()

    # 1) exact name match
    for c in candidates:
        c_stem = normalize_filename(Path(c["name"]).stem)
        c_ext = Path(c["name"]).suffix.lower()
        if c_stem == target_norm and (not ext or c_ext == ext):
            return c

    # 2) first candidate not excluded
    for c in candidates:
        c_stem = normalize_filename(Path(c["name"]).stem)
        if not should_exclude(c_stem, [s.lower() for s in exclusion_suffixes]):
            return c

    # 3) fallback to first candidate if any
    return candidates[0] if candidates else {}

def ensure_folder(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)

def download_file(service, file_id: str, file_name: str, company: str, base_folder: str) -> Path:
    """Download a single Drive file into base_folder/company/ and return local path."""
    company_dir = Path(base_folder) / company
    ensure_folder(company_dir)
    local_path = company_dir / file_name

    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(local_path, "wb")
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    return local_path

# ---------------------------
# Excel I/O
# ---------------------------

def read_input_excel(file) -> pd.DataFrame:
    """Read uploaded Excel file; must contain columns: filename, company."""
    df = pd.read_excel(file)
    # try to normalize expected columns
    cols = {c.strip().lower(): c for c in df.columns}
    if "filename" not in cols or "company" not in cols:
        raise ValueError("Excel must contain columns: 'filename' and 'company'")
    df = df.rename(columns={cols["filename"]: "filename", cols["company"]: "company"})
    df["filename"] = df["filename"].astype(str)
    df["company"] = df["company"].astype(str)
    return df[["filename", "company"]]

def write_results_excel(results: List[Dict], out_path: Path) -> None:
    """Write a workbook with one sheet per company listing Found/Not Found and links."""
    # group by company
    by_co: Dict[str, List[Dict]] = defaultdict(list)
    for row in results:
        by_co[row["company"]].append(row)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for company, rows in by_co.items():
            df = pd.DataFrame(rows)
            # reorder/report-friendly columns
            cols = ["company", "input_filename", "status", "file_name", "file_id", "webViewLink"]
            present = [c for c in cols if c in df.columns]
            df[present].to_excel(writer, index=False, sheet_name=company[:31])  # Excel sheet name limit

# ---------------------------
# Core Processing
# ---------------------------

def process_search(service, df: pd.DataFrame, exclusion_suffixes: List[str]) -> List[Dict]:
    """
    For each (filename, company), search Drive and produce a result row:
    {
      company, input_filename, status, file_name, file_id, webViewLink
    }
    """
    results = []
    progress = st.progress(0.0, text="Searching on Google Drive...")
    for i, row in df.iterrows():
        filename = str(row["filename"]).strip()
        company = str(row["company"]).strip()
        row_out = {
            "company": company,
            "input_filename": filename,
            "status": "Not Found",
            "file_name": "",
            "file_id": "",
            "webViewLink": ""
        }

        try:
            candidates = drive_search_by_name(service, filename)
            if candidates:
                match = best_match_from_candidates(filename, candidates, exclusion_suffixes)
                if match:
                    row_out.update({
                        "status": "Found",
                        "file_name": match.get("name", ""),
                        "file_id": match.get("id", ""),
                        "webViewLink": match.get("webViewLink", "")
                    })
        except Exception as e:
            row_out["status"] = f"Error: {e}"

        results.append(row_out)
        progress.progress((i + 1) / len(df))
    progress.empty()
    return results

def collect_download_list(results: List[Dict]) -> List[Dict]:
    """Return only the found files with required info for download."""
    return [
        {
            "company": r["company"],
            "file_id": r["file_id"],
            "file_name": r["file_name"]
        }
        for r in results
        if r.get("status") == "Found" and r.get("file_id")
    ]

# ---------------------------
# Streamlit UI
# ---------------------------

def ui():
    st.set_page_config(page_title="Drive File Search & Download", layout="centered")
    st.title("🔎 Google Drive File Search & Download")
    st.caption("Upload an Excel with `filename` and `company`. The app searches Drive, creates a per-company report, and can download the found files.")

    # Load config early to show defaults
    try:
        config = load_config()
    except Exception:
        st.error(
            "No `config.json` found. Create it in the repo root or set `CONFIG_PATH`. "
            "See README for a sample configuration."
        )
        st.stop()

    with st.expander("⚙️ Configuration (from config.json)", expanded=False):
        st.json(config)

    # Inputs
    st.subheader("1) Upload Excel input")
    uploaded = st.file_uploader("Excel (.xlsx) with columns: filename, company", type=["xlsx"])

    st.subheader("2) Output report")
    default_out = str(Path.cwd() / "output.xlsx")
    out_path_str = st.text_input("Results Excel file path", value=default_out)

    st.subheader("3) Download settings (optional)")
    download_folder = st.text_input("Base download folder", value=config.get("download_folder", "./downloads"))
    do_download = st.checkbox("Enable download of all found files", value=False)

    st.divider()
    run = st.button("▶️ Run Search")

    if not run:
        return

    # Validate inputs
    if not uploaded:
        st.warning("Please upload an input Excel file.")
        return

    try:
        df = read_input_excel(uploaded)
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
        return

    # Build Google service
    with st.status("Authorizing with Google...", expanded=False) as s:
        try:
            creds = get_credentials(
                token_file=config["token_file"],
                credentials_file=config["credentials_file"],
                scopes=config["scopes"]
            )
            service = build_drive_service(creds)
            s.update(label="Authorization complete ✅", state="complete")
        except Exception as e:
            s.update(label=f"Authorization failed: {e}", state="error")
            return

    # Run search
    with st.status("Searching files on Drive...", expanded=False) as s:
        results = process_search(service, df, config["exclusion_suffixes"])
        s.update(label="Search completed ✅", state="complete")

    # Summaries
    res_df = pd.DataFrame(results)
    found_count = (res_df["status"] == "Found").sum()
    not_found_count = (res_df["status"] != "Found").sum()

    st.subheader("Results Summary")
    col1, col2 = st.columns(2)
    col1.metric("Found", int(found_count))
    col2.metric("Not Found", int(not_found_count))
    st.dataframe(res_df, use_container_width=True)

    # Write Excel report
    out_path = Path(out_path_str)
    try:
        write_results_excel(results, out_path)
        st.success(f"Results workbook written to: {out_path}")
    except Exception as e:
        st.error(f"Failed to write results Excel: {e}")
        return

    # Optional download step
    if do_download and found_count > 0:
        if st.button("⬇️ Download All Found Files"):
            status_dl = st.empty()
            status_dl.info("Preparing to download files...")
            to_download = collect_download_list(results)

            # progress bar
            download_progress = st.progress(0.0)

            total = len(to_download)
            for i, file in enumerate(to_download):
                try:
                    download_file(
                        service,
                        file["file_id"],
                        file["file_name"],
                        file["company"],
                        download_folder
                    )
                except Exception as e:
                    st.warning(f"⚠️ Failed to download {file['file_name']}: {e}")
                download_progress.progress((i + 1) / total)

            status_dl.empty()
            st.success("✅ All files downloaded successfully into company folders.")

    st.info("Done. You can re-run with a different Excel or settings above.")

# ---------------------------
# Entry
# ---------------------------

if __name__ == "__main__":
    ui()
