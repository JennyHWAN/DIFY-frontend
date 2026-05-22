# AI-Driven SOC Report Generator

A web application that generates SOC 1 and SOC 2 audit reports from your control matrix files using a Dify AI workflow. Upload Excel, PDF, or Word files and receive a formatted `.docx` report.

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Running the App](#running-the-app)
- [Windows Shortcuts](#windows-shortcuts)
- [Building a Windows Executable](#building-a-windows-executable)
- [Using the App](#using-the-app)
- [Troubleshooting](#troubleshooting)

---

## Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| Python | 3.9 or higher | [Download here](https://www.python.org/downloads/) |
| pip | bundled with Python | Used to install dependencies |
| Dify API key | — | Required to run the workflow |

> **Windows users:** During Python installation, check **"Add Python to PATH"** or the app will not start.

---

## Installation

### Windows

```bat
pip install -r requirements.txt
```

### macOS / Linux

```bash
pip install -r requirements.txt
```

If you have both Python 2 and Python 3 installed on macOS/Linux, use `pip3` instead:

```bash
pip3 install -r requirements.txt
```

---

## Configuration

The app needs your Dify API key to call the workflow. You have two options:

### Option A — .env file (recommended for repeated use)

Create a file named `.env` in the same folder as `app.py`:

```
DIFY_API_BASE_URL=https://api.dify.ai/v1
DIFY_API_KEY=your_api_key_here
```

- `DIFY_API_BASE_URL` is optional — it defaults to `https://api.dify.ai/v1` if omitted.
- The `.env` file is loaded automatically when the app starts.
- Never commit this file to Git (it contains your secret key).

### Option B — Enter the key at runtime

Leave the sidebar **API Secret Key** field and type your key there each time you open the app. This overrides whatever is in `.env`.

---

## Running the App

### Windows

**Option 1 — Double-click (easiest):**
Double-click `run_app-win.bat`. It installs dependencies on the first run, then opens the app in your browser automatically.

**Option 2 — Command Prompt:**
```bat
python -m streamlit run app.py
```

### macOS

```bash
python3 -m streamlit run app.py
```

### Linux

```bash
python3 -m streamlit run app.py
```

The app opens at **http://localhost:8501** in your browser. If it does not open automatically, copy that URL and paste it into your browser.

To stop the app, press `Ctrl+C` in the terminal.

---

## Windows Shortcuts

Two `.bat` files are included for Windows users who prefer not to use the terminal:

| File | What it does |
|---|---|
| `run_app-win.bat` | Installs dependencies (first run only) and launches the app in your browser |
| `build_exe-win.bat` | Builds a standalone `.exe` that runs without Python installed |

---

## Building a Windows Executable

If you want to distribute the app to someone who does not have Python installed, you can build a self-contained `.exe` on Windows.

**Requirements:** Must be run on a Windows machine.

```bat
build_exe-win.bat
```

This will:
1. Install/upgrade all dependencies and PyInstaller
2. Clean any previous build output
3. Produce `dist\SOC_Report_Generator\SOC_Report_Generator.exe`

**Distributing the executable:**
1. Copy the entire `dist\SOC_Report_Generator\` folder to the target machine
2. In that folder, rename `.env.example` to `.env`
3. Fill in `DIFY_API_KEY` (and optionally `DIFY_API_BASE_URL`) in the `.env` file
4. Double-click `SOC_Report_Generator.exe`

> The `.exe` bundles Python and all libraries — no installation needed on the target machine.

---

## Using the App

### 1. Configure API (sidebar)

- The sidebar on the left shows **API Base URL** and **API Secret Key**
- If you set up a `.env` file the key is pre-loaded; otherwise enter it manually

### 2. Fill in Report Parameters

Required fields (marked with `*`):

| Field | Description |
|---|---|
| Company Name | Full legal name of the company |
| Company Short Name | Abbreviated name used in the filename |
| Service/System Name | Name of the system being audited |
| Service Description | Brief description of the service |
| Report Period Start | Start date, e.g. `2024-01-01` |
| Report Type | SOC1 TYPE1, SOC1 TYPE2, SOC2 TYPE1, or SOC2 TYPE2 |
| Output Language | English, 中文, or Both |
| Scope of Report | All carve out, Inclusive, or None |

Optional fields:

| Field | Description |
|---|---|
| Report Period End | Required for TYPE2 reports; enter `N/A` for TYPE1 |
| Company Website | Used in the report header |
| Industry | SaaS, Cloud Service, AI, PaaS, IaaS, General, or Other |
| Internal Supporting Systems | Auto-extracted from the control matrix if left empty |
| Control Domain | Only needed if not present in the uploaded document |
| Subservice Organization | Name of any subservice organization |

### 3. Trust Service Criteria (SOC 2 only)

Check the applicable criteria: Security, Availability, Processing Integrity, Confidentiality, Privacy.

### 4. Upload Control Matrix File(s)

Upload one or more files (Excel, PDF, Word). These are sent to the Dify workflow for analysis.

### 5. Generate Report

Click **Generate Report**. The workflow can take several minutes to complete.

When done:
- A preview of the result appears on screen
- Click **Download Report (.docx)** to save the Word document

The downloaded file is named: `{CompanyShortName}_{ReportType}_Report.docx`

---

## Troubleshooting

**"Python not found" on Windows**
- Reinstall Python from https://www.python.org/downloads/ and check **"Add Python to PATH"** during setup
- Or use the full path: `C:\Users\YourName\AppData\Local\Programs\Python\Python3x\python.exe`

**App opens but API key error appears**
- Check that your `.env` file is in the same folder as `app.py`
- Make sure the key does not have extra spaces or quotes around it

**Workflow times out**
- The workflow has a 600-second (10 minute) timeout. If it fails, try uploading a smaller file or splitting the control matrix

**Port 8501 already in use**
- Another instance of the app is running. Stop it with `Ctrl+C`, or run on a different port:
  ```bash
  python3 -m streamlit run app.py --server.port 8502
  ```

**macOS/Linux: `pip` installs but `streamlit` command not found**
- Use `python3 -m streamlit run app.py` instead of `streamlit run app.py`
- Or add your Python scripts folder to PATH: `export PATH="$HOME/.local/bin:$PATH"`
