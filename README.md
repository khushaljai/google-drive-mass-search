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

## ⚙️ Installation & Setup

### 1️⃣ Clone the repository
```bash
git clone https://github.com/yourusername/drive-file-search.git
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
