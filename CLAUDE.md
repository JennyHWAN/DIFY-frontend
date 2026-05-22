# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the App

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app runs on `http://localhost:8501` by default.

## Environment Variables

Create a `.env` file (loaded automatically via `python-dotenv`):

```
DIFY_API_BASE_URL=https://api.dify.ai/v1   # default if unset
DIFY_API_KEY=your_api_key_here
```

The API key can also be entered at runtime in the sidebar. The sidebar input overrides the env variable since `api_base` and `api_key_encrypted` are read from sidebar widgets; however, note that `api_key` is initialized from `API_KEY` and the sidebar currently only masks it — ensure the env var is set for non-interactive use.

## Building Windows Executable

```batch
build_exe-win.bat
```

This runs `pip install -r requirements.txt`, upgrades PyInstaller, and runs `pyinstaller app.spec`. The spec file is `app-win.spec`.

## Architecture

The entire application is a single file: `app.py`. There is no backend server — the Streamlit app communicates directly with the Dify API.

**Data flow:**
1. User fills in report parameters and uploads control matrix files (Excel, PDF, Word)
2. Files are POSTed to Dify's `/files/upload` endpoint; each returns an `upload_file_id`
3. All form inputs plus the file IDs are sent as `inputs` to `/workflows/run` (blocking mode, 600s timeout)
4. The workflow returns a markdown string under `outputs.Result`
5. `markdown_to_docx()` converts the markdown to a `.docx` (python-docx), handling headings H1–H4, bullet/numbered lists, bold, and italic
6. User downloads the `.docx`

**Key functions in `app.py`:**
- `upload_file_to_dify()` — multipart POST to Dify files API
- `run_workflow()` — blocking POST to Dify workflows API; extracts `data.outputs`
- `markdown_to_docx()` — line-by-line markdown parser to python-docx Document
- `_add_inline_formatting()` — regex-based `**bold**` / `*italic*` inline parser

**Dify workflow definition:** `AI-driven report generation-complete.yml` — the full Dify workflow YAML (7,896 lines). This is imported into the Dify platform, not executed locally.

## Workflow Input Variables

The workflow expects these input keys (see `inputs` dict in `app.py`):

| Key | Description |
|-----|-------------|
| `File_input` | List of `{transfer_method, upload_file_id, type}` dicts |
| `Report_type` | `"SOC1 TYPE1"`, `"SOC1 TYPE2"`, `"SOC2 TYPE1"`, `"SOC2 TYPE2"` |
| `Output_language` | `"English"`, `"中文"`, `"Both"` |
| `is_Security`, `is_Availability`, `is_Processing_Integrity`, `is_Confidentiality`, `is_Privacy` | Boolean TSC checkboxes (SOC2 only) |
| Company/service fields | `Company_name`, `Co_short_name`, `Industry`, `System_or_service_name`, etc. |
