# AI-Driven SOC Report Generator

A Streamlit application that generates SOC 1 and SOC 2 audit reports from control matrix files. Upload Excel, PDF, or Word files and receive a formatted `.docx` report.

The final document combines two pipelines:

1. **EY Word templates** (local, no AI) — Section I (Management Assertion) and Section II (Independent Auditor's Report), filled in from the form fields.
2. **Dify AI workflow** (three streaming steps) — Sections III–IV, the generated report body.

Both are assembled into a single `.docx`. You can also generate just the template sections, or just the Dify body.

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

Everything is on a single page. Fill in the form, then click **Run All Steps** — the app uploads the files, runs the three Dify workflow steps (MAIN → SUB1 → SUB2) back-to-back, assembles the document, and offers it for download. There are no separate per-step buttons; progress is shown per node as it runs.

### Complete Report Settings

**Generate complete report (MA + AR + main sections)** is checked by default. When enabled, the output includes the two EY-template sections ahead of the Dify-generated body:

- **Section I — Management Assertion (MA)**
- **Section II — Independent Auditor's Report (AR)**

Uncheck it to generate only the Dify sections (III–IV).

When enabled, the **Complete Report Settings** panel drives the templates:

| Setting | Description |
|---|---|
| Standard | Audit standard (the available options depend on Report Type) |
| Report Signing Date | `YYYY-MM-DD`; auto-formatted as "January 30, 2026" (EN) / "2026年1月30日" (CN) |
| Signing City | City shown in the signature block, e.g. Shanghai |
| AR Addressee | "Management" or "Board of Directors" — sets the AR opening line |
| Complementary User Entity Controls (CUEC) | Identified / Not Identified — controls CUEC wording in the AR |
| Includes transaction processing wording | When unchecked, the **Systems Function** field is used in place of the transaction-processing wording |
| Single user entity report | Toggles single-user-entity template spans |
| Subject matter includes AI technology | Adds the paragraph noting AI is used but excluded from audit scope |
| Report includes 'Other Information' section | When unchecked, removes paragraphs that reference the Other Information section |
| SSO Complementary Controls | Identified / Not Identified (shown only when a Subservice Organization testing strategy other than *None* is selected) |

The panel also shows a **live template-resolution preview** indicating which MA/AR `.docx` will be used for the chosen Report Type / Standard / SSO / Language combination, or a warning if none matches.

#### How Section I & II templates are chosen

Section I (Management Assertion) and Section II (Independent Auditor's Report) are **not** AI-generated — each is filled from a pre-authored EY Word template. The app picks the right `.docx` per section by looking your form selections up in `template_index.xlsx`, then matching the resulting work-paper number to a file on disk.

**1. Your UI choices are mapped to the spreadsheet's vocabulary** (`resolve_template()` in `app.py`):

| Form selection | Lookup column | Mapped value |
|---|---|---|
| Report Type | `Category` | `SOC1 *` → `SOC 1`, `SOC2 *` → `SOC 2` |
| Report Type | `Type` | `* TYPE1` → `Type I`, `* TYPE2` → `Type II` |
| Standard | `Standards` | passed through (`SSAE 18`, `ISAE 3402`, `ISAE 3000`); any "Combined" choice → `Combined` |
| Subservice Org Testing Strategy | `Sub-service Organization (SSO)` | `None` → `none`, `All carve out` → `all carve out`, `Inclusive` → `Inclusive` |
| Output Language | `Language` | `English` → `EN`, `中文` → `CN` |

**2. The spreadsheet returns a work-paper number.** `template_index.xlsx` has one sheet per section — **`MA`** and **`AR`** — each row being a unique combination of the five columns above plus a **WP number** (e.g. `1.1`, `13.2`). The app scans the matching sheet for the row where all five mapped values match and reads its WP number. (The `.1`/`.2` suffix tracks EN/CN.)

**3. The WP number selects the actual Word file.** The app lists `MA_template/` (for Section I) or `AR_template/` (for Section II) and picks the `.docx` whose filename **starts with that WP number followed by a space** — so WP `13.1` resolves to `AR_template/13.1 AR_SOC2 Type II_SSAE18_IL503_EN（none SSO）.docx`. The rest of the filename is just a human-readable description; only the leading number is matched.

**Worked example** — _SOC2 TYPE2 · SSAE 18 · None SSO · English_ resolves to:

- Section I → MA sheet → WP `11.1` → `MA_template/11.1 MA_SOC2 Type II_IL508_EN（none SSO）.docx`
- Section II → AR sheet → WP `13.1` → `AR_template/13.1 AR_SOC2 Type II_SSAE18_IL503_EN（none SSO）.docx`

**When no template matches:** if the combination has no row (or the row's WP cell is blank, e.g. an `N/A`/"No template" entry), the live preview shows a warning and that section is skipped. Not every Report Type × Standard × SSO × Language combination has an authored template — the MA and AR sheets define exactly which ones do. To support a new combination, add the template `.docx` to the right folder (prefixed with a new WP number) and add the matching row to the spreadsheet.

### Upload & form fields

**Upload:** One or more Excel, PDF, or Word files containing the control matrix and supporting data.

**Required fields:**

| Field | Description |
|---|---|
| Company Name | Full legal name |
| Company Short Name | Abbreviated name (used in the output filename) |
| Service / System Name | Name of the system being audited |
| Service Description | Brief description of the service |
| Report Period Start | Start date, e.g. `2025-01-01` (use as-of date for TYPE1) |
| Report Period End | Required for TYPE2; leave blank (N/A) for TYPE1 |
| Report Type | SOC1 TYPE1, SOC1 TYPE2, SOC2 TYPE1, or SOC2 TYPE2 |
| Output Language | English or 中文 |
| Subservice Organization Testing Strategy | None, All carve out, or Inclusive |
| Subservice Organization | Name and services (one per line: `Org Name \| Services`); required if testing strategy is not *None* |
| Industry | HR, IaaS, AI, SaaS, or Others |

**Optional fields:**

| Field | Description |
|---|---|
| Control Domain | Only needed if not present in the uploaded document |
| Company Website | Included in the report header |
| Systems Function | Describes the purpose of the internal supporting systems |
| Internal Supporting Systems | Auto-extracted from the control matrix if left blank |

**Trust Service Criteria (SOC 2 only):** Check Security, Availability, Processing Integrity, Confidentiality, and/or Privacy as applicable. At least one is required for SOC 2 reports.

**User Entity Section** — two toggles for optional report sections generated from the control matrix:

- **Include CUEC** (Complementary User Entity Controls) — on by default for SOC1 reports
- **Include UER** (User Entity Responsibilities) — on by default for SOC2 reports

The default follows the Report Type but can be overridden within a given type.

### Generating the report

- **▶ Run All Steps (1 → 2 → 3)** — uploads files and runs the full Dify pipeline, then assembles and offers the final `.docx`. Typically takes several minutes.
- **🧪 Generate MA + AR only (templates, no Dify)** — fills the Section I + II templates from the form fields *without* calling Dify (handy for checking template output). Produces a `{ShortName}_{ReportType}_MA_AR.docx` download.

### Download

When the run finishes, a report preview appears and a **Download Report (.docx)** button is shown.

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
