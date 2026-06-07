# google-drive-mass-search
Streamlit app that automates Google Drive file lookup and organization using an Excel reference list. Upload a sheet with filename and company, search Drive securely via OAuth, generate Found/Not Found reports, and optionally download matched files into company-wise folders.
# Google Drive File Search & Download (Streamlit)

A Streamlit-based Python app that automates Google Drive file lookup, reporting, and downloading using an Excel reference list.  
Perfect for finance, audit, and compliance professionals who need to retrieve and organize large sets of Drive files efficiently.

---

## 🚀 Features

- 🔍 **Search Google Drive** for filenames from an Excel file  
- 🏢 **Group results by company** (based on your Excel input)  
- 📊 **Generate a structured Excel report** showing “Found” and “Not Found” files  
- ⬇️ **Download all matched files** into company-specific folders  
- ⚙️ **Configurable behavior** via `config.json`  
- 🔒 **Secure OAuth 2.0 authentication** (no API keys stored in code)

---

## 📂 Project Structure

📁 your-repo/
┣ 📄 appv7.py
┣ 📄 config.json
┣ 📄 requirements.txt
┣ 📄 README.md
┣ 📁 downloads/
┗ 📄 credentials.json (Google OAuth client credentials)

yaml
Copy code

---

## 🧰 Requirements

- Python 3.9 or higher  
- Google Cloud project with **Drive API** enabled  
- A Google OAuth **Desktop app** client ID (JSON credentials)

---

## 🔐 Security & Privacy

No hardcoded API keys are stored.

OAuth token (token.pickle) is created locally on first run.

Keep credentials.json, token.pickle, and all personal files out of version control.

Use .gitignore to ensure these files remain private.

## 🧩 How It Works (Under the Hood)

Loads configuration and authenticates via Google OAuth.

Reads the uploaded Excel file using pandas.

For each filename, uses the Google Drive API (files.list) with a name contains query.

Filters out duplicates and excluded suffixes.

Writes all results into an Excel report, grouped per company.

Downloads found files into structured folders if requested.

## 🧑‍💻 Technologies Used

Python

Streamlit (UI)

Pandas (Excel I/O)

Google Drive API

OAuth 2.0 for authentication

OpenPyXL for Excel writing

## 🏁 Summary

This app provides a simple and secure workflow to:

Search Google Drive in bulk

Organize results by company

Generate clear reports

Download files efficiently

No credentials are exposed, and everything runs locally.

## ▶️ Running the App

Start the Streamlit app:

streamlit run appv7.py

## Usage of the app
Once it opens in your browser:

Upload Excel file — must contain columns filename and company.

Specify an output Excel file path (for the results workbook).

Define the download folder (optional).

Click Run Search to begin Drive scanning.

After results are generated, click Download All Found Files to save them locally.


## ⚡ Troubleshooting
Issue	Solution
Auth window doesn’t open	Ensure your OAuth client type is Desktop. Try deleting token.pickle and reauthorizing.
403 / Insufficient permissions	Recreate credentials with drive.readonly scope enabled.
Excel read error	Make sure your file is .xlsx and has columns filename and company.
No results found	Check Drive sharing permissions and verify filenames (case-sensitive).
Download fails	Ensure the target folder is writable and not locked by another process.

## ⚙️ Installation & Setup

### 1️⃣ Clone the repository
```bash
git clone https://github.com/khushaljai/drive-file-search.git
cd drive-file-search
2️⃣ Create and activate a virtual environment
bash
Copy code
python -m venv .venv
# On Windows
.venv\Scripts\activate
# On macOS/Linux
source .venv/bin/activate
3️⃣ Install dependencies
bash
Copy code
pip install -r requirements.txt
4️⃣ Enable the Google Drive API
Go to Google Cloud Console.

Create a project (if you don’t have one already).

Enable the Google Drive API.

Create OAuth 2.0 Client ID → choose Desktop app.

Download the credentials file and rename it to credentials.json.

Place it in the same folder as appv7.py.



