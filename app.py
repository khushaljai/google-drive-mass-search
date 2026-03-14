#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Google Drive File Search & Download (Streamlit)

Features:
- Upload an Excel with columns: filename, company
- Search Google Drive for each filename
- Generate per-company "Found / Not Found" results workbook
- Optionally download matched files into company subfolders
- Supports downloading both:
  - regular Drive files via get_media()
  - Google Docs/Sheets/Slides via export_media()

Config:
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

import io
import json
import os
import pickle
import re
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import streamlit as st
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload


# ---------------------------
# Config Helpers
# ---------------------------

def load_config() -> Dict:
    """Load configuration from CONFIG_PATH env var or local config.json."""
    cfg_path = Path(os.getenv("CONFIG_PATH", "config.json"))
    if not cfg_path.exists():
        raise FileNotFoundError(
            f"Config file not found: {cfg_path}. Create config.json or set CONFIG_PATH."
        )

    with cfg_path.open("r", encoding="utf-8") as f:
        cfg = json.load(f)

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
    """Return Google Credentials object using token file and credentials file."""
    creds = None
    token_path = Path(token_file)

    if token_path.exists():
        try:
            with token_path.open("rb") as token:
                creds = pickle.load(token)
        except Exception:
            creds = None

    if creds and getattr(creds, "expired", False) and getattr(creds, "refresh_token", None):
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
# Utility Helpers
# ---------------------------

def ensure_folder(path: Path) -> None:
    """Create folder if it does not exist."""
    path.mkdir(parents=True, exist_ok=True)


def normalize_filename(name: str) -> str:
    """
    Normalize filename stem for matching:
    - remove trailing counters like ' (1)'
    - strip spaces
    - lowercase
    """
    if not isinstance(name, str):
        return ""
    name = re.sub(r"\s*\(\d+\)\s*$", "", name.strip())
    return name.strip().lower()


def should_exclude(name_no_ext: str, exclusion_suffixes: List[str]) -> bool:
    """Return True if the stem ends with any excluded suffix."""
    stem = (name_no_ext or "").lower()
    return any(stem.endswith(suf.lower()) for suf in exclusion_suffixes)


def sanitize_sheet_name(name: str) -> str:
    """Make a safe Excel sheet name."""
    invalid = r'[]:*?/\\'
    cleaned = "".join("_" if c in invalid else c for c in str(name))
    cleaned = cleaned.strip() or "Sheet"
    return cleaned[:31]


def sanitize_folder_name(name: str) -> str:
    """Make a safe local folder name."""
    return re.sub(r'[<>:"/\\|?*]+', "_", str(name)).strip() or "Unknown"


# ---------------------------
# Drive Search Helpers
# ---------------------------

def drive_search_by_name(service, filename: str, max_results: int = 20) -> List[Dict]:
    """
    Search Drive files by name.
    Uses the stem if extension is present so matching is broader.
    """
    filename = filename.strip()
    stem = Path(filename).stem if Path(filename).suffix else filename
    safe = stem.replace('"', '\\"')

    query = f'name contains "{safe}" and trashed = false'
    results = service.files().list(
        q=query,
        spaces="drive",
        fields="files(id, name, mimeType, size, webViewLink)",
        pageSize=max_results
    ).execute()

    return results.get("files", [])


def best_match_from_candidates(
    filename: str,
    candidates: List[Dict],
    exclusion_suffixes: List[str]
) -> Optional[Dict]:
    """
    Select best match from candidates:
    1. exact stem + extension match
    2. exact stem match
    3. first non-excluded candidate
    4. first candidate
    """
    if not candidates:
        return None

    target_stem = normalize_filename(Path(filename).stem)
    target_ext = Path(filename).suffix.lower()

    # 1) Exact stem + extension match
    for c in candidates:
        c_name = c.get("name", "")
        c_stem = normalize_filename(Path(c_name).stem)
        c_ext = Path(c_name).suffix.lower()
        if c_stem == target_stem and c_ext == target_ext:
            return c

    # 2) Exact stem match
    for c in candidates:
        c_name = c.get("name", "")
        c_stem = normalize_filename(Path(c_name).stem)
        if c_stem == target_stem:
            return c

    # 3) First non-excluded candidate
    for c in candidates:
        c_name = c.get("name", "")
        c_stem = normalize_filename(Path(c_name).stem)
        if not should_exclude(c_stem, exclusion_suffixes):
            return c

    # 4) Fallback
    return candidates[0]


# ---------------------------
# Download Helpers
# ---------------------------

EXPORT_MIME_MAP = {
    "application/vnd.google-apps.document": (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".docx",
    ),
    "application/vnd.google-apps.spreadsheet": (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ".xlsx",
    ),
    "application/vnd.google-apps.presentation": (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        ".pptx",
    ),
    "application/vnd.google-apps.drawing": (
        "image/png",
        ".png",
    ),
}


def download_file(
    service,
    file_id: str,
    file_name: str,
    company: str,
    base_folder: str,
    mime_type: str
) -> Path:
    """
    Download one file into base_folder/company/.
    Uses export_media for Google-native files and get_media for regular files.
    """
    company_dir = Path(base_folder) / sanitize_folder_name(company)
    ensure_folder(company_dir)

    if mime_type in EXPORT_MIME_MAP:
        export_mime, default_ext = EXPORT_MIME_MAP[mime_type]
        local_name = file_name if Path(file_name).suffix else f"{file_name}{default_ext}"
        local_path = company_dir / local_name
        request = service.files().export_media(fileId=file_id, mimeType=export_mime)
    else:
        local_path = company_dir / file_name
        request = service.files().get_media(fileId=file_id)

    with io.FileIO(local_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()

    return local_path


def collect_download_list(results: List[Dict]) -> List[Dict]:
    """Return only found rows needed for downloading."""
    return [
        {
            "company": r["company"],
            "file_id": r["file_id"],
            "file_name": r["file_name"],
            "mimeType": r.get("mimeType", ""),
        }
        for r in results
        if r.get("status") == "Found" and r.get("file_id")
    ]


# ---------------------------
# Excel I/O
# ---------------------------

def read_input_excel(file) -> pd.DataFrame:
    """Read uploaded Excel file; it must contain columns: filename, company."""
    df = pd.read_excel(file)

    cols = {str(c).strip().lower(): c for c in df.columns}
    if "filename" not in cols or "company" not in cols:
        raise ValueError("Excel must contain columns: 'filename' and 'company'")

    df = df.rename(columns={
        cols["filename"]: "filename",
        cols["company"]: "company"
    })

    df["filename"] = df["filename"].astype(str).str.strip()
    df["company"] = df["company"].astype(str).str.strip()

    # remove empty rows
    df = df[(df["filename"] != "") & (df["company"] != "")]
    return df[["filename", "company"]]


def write_results_excel(results: List[Dict], out_path: Path) -> None:
    """Write results workbook with one sheet per company."""
    by_company: Dict[str, List[Dict]] = defaultdict(list)
    for row in results:
        by_company[row["company"]].append(row)

    ensure_folder(out_path.parent)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for company, rows in by_company.items():
            df = pd.DataFrame(rows)

            preferred_cols = [
                "company",
                "input_filename",
                "status",
                "file_name",
                "file_id",
                "mimeType",
                "webViewLink",
                "error_message",
            ]
            cols = [c for c in preferred_cols if c in df.columns]
            df[cols].to_excel(
                writer,
                index=False,
                sheet_name=sanitize_sheet_name(company)
            )


# ---------------------------
# Core Processing
# ---------------------------

def process_search(service, df: pd.DataFrame, exclusion_suffixes: List[str]) -> List[Dict]:
    """Search Drive for each filename and return row-wise results."""
    results: List[Dict] = []

    progress_bar = st.progress(0.0, text="Searching on Google Drive...")

    total = len(df)
    for i, row in enumerate(df.itertuples(index=False), start=1):
        filename = str(row.filename).strip()
        company = str(row.company).strip()

        row_out = {
            "company": company,
            "input_filename": filename,
            "status": "Not Found",
            "file_name": "",
            "file_id": "",
            "mimeType": "",
            "webViewLink": "",
            "error_message": "",
        }

        try:
            candidates = drive_search_by_name(service, filename)
            match = best_match_from_candidates(filename, candidates, exclusion_suffixes)

            if match:
                row_out.update({
                    "status": "Found",
                    "file_name": match.get("name", ""),
                    "file_id": match.get("id", ""),
                    "mimeType": match.get("mimeType", ""),
                    "webViewLink": match.get("webViewLink", ""),
                })

        except Exception as e:
            row_out["status"] = "Error"
            row_out["error_message"] = str(e)

        results.append(row_out)
        progress_bar.progress(i / total, text=f"Searching on Google Drive... ({i}/{total})")

    progress_bar.empty()
    return results


# ---------------------------
# Session State
# ---------------------------

def init_session_state():
    """Initialize Streamlit session state keys."""
    defaults = {
        "service": None,
        "results": None,
        "config": None,
        "input_df": None,
        "last_output_path": None,
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# ---------------------------
# UI
# ---------------------------

def ui():
    st.set_page_config(page_title="Drive File Search & Download", layout="centered")
    init_session_state()

    st.title("🔎 Google Drive File Search & Download")
    st.caption(
        "Upload an Excel with `filename` and `company`. "
        "The app searches Drive, creates a per-company report, and can download the found files."
    )

    try:
        config = load_config()
        st.session_state.config = config
    except Exception as e:
        st.error(str(e))
        st.stop()

    with st.expander("⚙️ Configuration (from config.json)", expanded=False):
        st.json(st.session_state.config)

    st.subheader("1) Upload Excel input")
    uploaded = st.file_uploader(
        "Excel (.xlsx) with columns: filename, company",
        type=["xlsx"]
    )

    st.subheader("2) Output report")
    default_out = str(Path.cwd() / "output.xlsx")
    out_path_str = st.text_input(
        "Results Excel file path",
        value=st.session_state.last_output_path or default_out
    )

    st.subheader("3) Download settings")
    download_folder = st.text_input(
        "Base download folder",
        value=st.session_state.config.get("download_folder", "./downloads")
    )
    do_download = st.checkbox("Enable download of all found files", value=False)

    st.divider()

    run_search = st.button("▶️ Run Search")

    if run_search:
        if not uploaded:
            st.warning("Please upload an input Excel file.")
            st.stop()

        try:
            df = read_input_excel(uploaded)
            st.session_state.input_df = df
        except Exception as e:
            st.error(f"Failed to read Excel: {e}")
            st.stop()

        with st.status("Authorizing with Google...", expanded=False) as status_box:
            try:
                creds = get_credentials(
                    token_file=st.session_state.config["token_file"],
                    credentials_file=st.session_state.config["credentials_file"],
                    scopes=st.session_state.config["scopes"],
                )
                service = build_drive_service(creds)
                st.session_state.service = service
                status_box.update(label="Authorization complete ✅", state="complete")
            except Exception as e:
                status_box.update(label=f"Authorization failed: {e}", state="error")
                st.stop()

        with st.status("Searching files on Drive...", expanded=False) as status_box:
            try:
                results = process_search(
                    st.session_state.service,
                    st.session_state.input_df,
                    st.session_state.config["exclusion_suffixes"],
                )
                st.session_state.results = results
                st.session_state.last_output_path = out_path_str
                status_box.update(label="Search completed ✅", state="complete")
            except Exception as e:
                status_box.update(label=f"Search failed: {e}", state="error")
                st.stop()

        try:
            write_results_excel(st.session_state.results, Path(out_path_str))
            st.success(f"Results workbook written to: {out_path_str}")
        except Exception as e:
            st.error(f"Failed to write results Excel: {e}")

    # Show existing results after reruns
    if st.session_state.results:
        res_df = pd.DataFrame(st.session_state.results)

        found_count = int((res_df["status"] == "Found").sum())
        not_found_count = int((res_df["status"] == "Not Found").sum())
        error_count = int((res_df["status"] == "Error").sum())

        st.subheader("Results Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("Found", found_count)
        col2.metric("Not Found", not_found_count)
        col3.metric("Errors", error_count)

        st.dataframe(res_df, use_container_width=True)

        if do_download and found_count > 0:
            if st.button("⬇️ Download All Found Files"):
                if st.session_state.service is None:
                    st.error("Google Drive service is not available. Please run the search again.")
                    st.stop()

                to_download = collect_download_list(st.session_state.results)
                total = len(to_download)

                if total == 0:
                    st.warning("No downloadable files found.")
                    st.stop()

                st.info("Starting file download...")
                progress_bar = st.progress(0.0)
                downloaded = 0
                failed = 0

                for i, file in enumerate(to_download, start=1):
                    try:
                        download_file(
                            service=st.session_state.service,
                            file_id=file["file_id"],
                            file_name=file["file_name"],
                            company=file["company"],
                            base_folder=download_folder,
                            mime_type=file["mimeType"],
                        )
                        downloaded += 1
                    except Exception as e:
                        failed += 1
                        st.warning(f"Failed to download '{file['file_name']}': {e}")

                    progress_bar.progress(i / total)

                st.success(
                    f"Download completed. Success: {downloaded}, Failed: {failed}. "
                    f"Files saved in: {download_folder}"
                )
        elif do_download and found_count == 0:
            st.info("No found files available for download.")

    st.info("Done. You can upload another file or rerun the search with different settings.")


# ---------------------------
# Entry
# ---------------------------

if __name__ == "__main__":
    ui()
