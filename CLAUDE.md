# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

See also: `../CLAUDE.md` at the repository root for the full architecture overview.

## Running the App

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

The app runs on `http://localhost:8501` by default.

## Environment Variables

Create a `.env` file (use `.env.example` as a template):

```
DIFY_API_BASE_URL=https://api.dify.ai/v1
DIFY_API_KEY_MAIN=app-...   # Step 1 — MAIN workflow
DIFY_API_KEY_SUB1=app-...   # Step 2 — SUB1 workflow
DIFY_API_KEY_SUB2=app-...   # Step 3 — SUB2 workflow
```

## Building Windows Executable

```batch
build_exe-win.bat
```

Requires `.env` with valid keys to exist before building — the keys are compiled into `_bundled_config.py` (bytecode) and bundled into the exe so end-users need no configuration. Output goes to `C:\DIFY_build\dist\SOC_Report_Generator\`.

## Architecture

The entire application is a single file: `app.py`. There is no backend server — the Streamlit app communicates directly with the Dify API using **streaming (SSE)** with a 30s connect / 1800s read timeout.

**Key functions in `app.py`:**

| Function | Purpose |
|---|---|
| `upload_file(file_bytes, filename, api_base, api_key)` | Multipart POST to `/files/upload`; returns `upload_file_id`. Uses `verify=False` (intentional for internal use). |
| `run_workflow(inputs, api_base, api_key, status_placeholder=None)` | Streaming POST to `/workflows/run`; parses SSE events internally, updates `status_placeholder` on `node_finished` events, returns `outputs` dict on `workflow_finished`. |
| `markdown_to_docx(md_text)` | Converts markdown to `.docx` bytes (H1–H4, bullet/numbered lists, bold, italic). Returns `bytes` via `io.BytesIO`, not a `Document`. |
| `_inline(para, text)` | Regex-based `**bold**` / `*italic*` inline parser. Does not handle nested or escaped asterisks. |
| `_apply_fonts(run)` | Sets Times New Roman (Latin) + 黑体 (CJK) dual fonts on a run. |
| `to_str(v)` | Coerces any workflow output value to `str` before passing to next step. `None` → `""`, lists/dicts → JSON with `ensure_ascii=False`. |

**Three-step data flow:**

1. User uploads files + fills form → Step 1 (MAIN) → `main_outputs` stored in session state
2. Step 2 (SUB1) reads `main_outputs` → `sub1_outputs` stored in session state
3. Step 3 (SUB2) reads both → `final_result` (markdown string) stored in session state
4. `markdown_to_docx()` converts `final_result` → `.docx` download named `{Co_short_name}_{Report_type}_Report.docx` (spaces → underscores)

The sidebar **Reset** button clears all four session state keys: `main_outputs`, `sub1_outputs`, `final_result`, `user_inputs`.

**Error handling:** Dify workflows may return graceful error fields instead of HTTP errors. After each step, `outputs.get("Error")` and `outputs.get("Error_ORG")` are checked and surfaced to the user before stopping.

**SUB2 fallback pattern:** SUB2 uses `to_str(sub1_outputs.get(key) or user_inputs.get(key))` for fields that SUB1 may not modify, ensuring continuity even if a field was not present in SUB1's outputs.

## Workflow Input Variables (Step 1 — MAIN)

| Key | Description |
|---|---|
| `File_input` | List of `{transfer_method, upload_file_id, type}` dicts |
| `Report_type` | `"SOC1 TYPE1"`, `"SOC1 TYPE2"`, `"SOC2 TYPE1"`, `"SOC2 TYPE2"` |
| `Output_language` | `"English"`, `"中文"`, `"Both"` |
| `is_Security`, `is_Availability`, `is_Processing_Integrity`, `is_Confidentiality`, `is_Privacy` | Boolean TSC flags (SOC2 only; sent for all report types but only used by Dify when applicable) |
| `Company_name`, `Co_short_name`, `Industry`, `System_or_service_name`, … | Company/service metadata fields |
