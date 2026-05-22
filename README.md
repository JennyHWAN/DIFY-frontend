# AI-Driven SOC Report Generator

A Streamlit application that generates SOC 1 and SOC 2 audit reports from control matrix files using a Dify AI workflow. Upload Excel, PDF, or Word files and receive a formatted `.docx` report.

---

## Table of Contents

- [Prerequisites](#prerequisites)
- [Developer Setup](#developer-setup)
- [Configuration](#configuration)
- [Running the App Locally](#running-the-app-locally)
- [Building the Windows Executable](#building-the-windows-executable)
- [Distributing to Users](#distributing-to-users)
- [Using the App](#using-the-app)
- [Troubleshooting](#troubleshooting)

---

## Prerequisites

| Requirement | Version | Notes |
|---|---|---|
| Python | 3.9 or higher | [Download here](https://www.python.org/downloads/) |
| pip | bundled with Python | Used to install dependencies |
| Dify API keys | — | Three keys required (MAIN, SUB1, SUB2 workflows) |

> **Windows users:** During Python installation, check **"Add Python to PATH"**.

---

## Developer Setup

```bat
pip install -r requirements.txt
```

---

## Configuration

API keys are loaded from a `.env` file. This file is gitignored and must never be committed.

Create `.env` in the project root (use `.env.example` as a template):

```
DIFY_API_BASE_URL=https://api.dify.ai/v1
DIFY_API_KEY_MAIN=app-xxxxxxxxxxxxxxxx
DIFY_API_KEY_SUB1=app-xxxxxxxxxxxxxxxx
DIFY_API_KEY_SUB2=app-xxxxxxxxxxxxxxxx
```

- `DIFY_API_KEY_MAIN` — used by Step 1 (Extract & Prepare)
- `DIFY_API_KEY_SUB1` — used by Step 2 (Entity Level Controls)
- `DIFY_API_KEY_SUB2` — used by Step 3 (Final Report Assembly)

The `.env` file is also **baked into the Windows executable** at build time so end users require no configuration.

---

## Running the App Locally

```bat
python -m streamlit run app.py
```

The app opens at **http://localhost:8501**. Press `Ctrl+C` in the terminal to stop it.

---

## Building the Windows Executable

Run on a Windows machine. The `.env` file must exist and contain valid keys before building — it is bundled into the output.

```bat
build_exe-win.bat
```

The script will:
1. Install / upgrade all dependencies and PyInstaller
2. Clean previous build output from `C:\DIFY_build\`
3. Bundle the app, all libraries, and the `.env` into `C:\DIFY_build\dist\SOC_Report_Generator\`

> Build output is written to `C:\DIFY_build\` (not inside the project folder) to avoid Windows MAX_PATH errors caused by the long paths Streamlit creates inside the bundle.

---

## Distributing to Users

1. Copy the entire `C:\DIFY_build\dist\SOC_Report_Generator\` folder to the target machine
2. Users double-click `SOC_Report_Generator.exe` — no Python, no configuration required

API keys are embedded inside `_internal\` within the bundle and are not visible to end users.

---

## Using the App

The app runs as a three-step workflow. Each step must complete before the next becomes available.

### Step 1 — MAIN: Extract & Prepare

**Upload files:** One or more Excel, PDF, or Word files containing the control matrix and supporting data.

**Required fields:**

| Field | Description |
|---|---|
| Company Name | Full legal name |
| Company Short Name | Abbreviated name (used in the output filename) |
| Service / System Name | Name of the system being audited |
| Service Description | Brief description of the service |
| Report Period Start | Start date, e.g. `2025-01-01` (use as-of date for TYPE1) |
| Report Type | SOC1 TYPE1, SOC1 TYPE2, SOC2 TYPE1, or SOC2 TYPE2 |
| Output Language | English or 中文 |
| Subservice Organization Testing Strategy | None, All carve out, or Inclusive |

**Optional fields:**

| Field | Description |
|---|---|
| Report Period End | Required for TYPE2; leave blank for TYPE1 |
| Company Website | Included in the report header |
| Industry | SaaS, Cloud Service, AI, PaaS, IaaS, General, or Other |
| Internal Supporting Systems | Auto-extracted from the control matrix if left blank |
| Systems Function | Describes the purpose of the supporting systems |
| Control Domain | Only needed if not present in the uploaded document |
| Subservice Organization | Name and services; required if testing strategy is not None |

**Trust Service Criteria (SOC 2 only):** Check Security, Availability, Processing Integrity, Confidentiality, and/or Privacy as applicable.

Click **Run Step 1 — MAIN Workflow** to start. This step typically takes several minutes.

### Step 2 — SUB1: Entity Level Controls

Appears after Step 1 completes. Click **Run Step 2 — SUB1 Workflow**. No additional inputs required.

### Step 3 — SUB2: Final Report Assembly

Appears after Step 2 completes. Click **Run Step 3 — SUB2 Workflow**. No additional inputs required.

### Download

When Step 3 finishes, a report preview appears and a **Download Report (.docx)** button is shown.

The downloaded file is named: `{CompanyShortName}_{ReportType}_Report.docx`

---

## Troubleshooting

**Executable flashes and closes immediately**
- A `launch_error.log` file is created next to `SOC_Report_Generator.exe` — open it to see the full error.

**"MAIN/SUB1/SUB2 API key is not set" error**
- The `.env` was missing when the executable was built. Add the keys to `.env` and rebuild.

**Workflow times out or returns an error**
- The workflow uses a 30-minute streaming timeout. Try uploading a smaller or simpler control matrix file.
- Check that the Dify platform is reachable from the machine running the exe.

**Port 8501 already in use (local dev only)**
- Another instance is running. Stop it with `Ctrl+C`, or run on a different port:
  ```bat
  python -m streamlit run app.py --server.port 8502
  ```
