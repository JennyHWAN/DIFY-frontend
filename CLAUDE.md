# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

See also: `../CLAUDE.md` at the workspace root for how this frontend connects to the `DIFY-backend` workflows.

## Running the App

```bash
pip install -r requirements.txt
python -m streamlit run app.py            # http://localhost:8501
python -m streamlit run app.py --server.port 8502   # if 8501 is taken
```

There is no test suite or linter — verification is manual through the running app.

## Environment Variables

Create a `.env` file (use `.env.example` as a template):

```
DIFY_API_BASE_URL=https://api.dify.ai/v1
DIFY_API_KEY_MAIN=app-...   # Step 1 — MAIN workflow
DIFY_API_KEY_SUB1=app-...   # Step 2 — SUB1 workflow
DIFY_API_KEY_SUB2=app-...   # Step 3 — SUB2 workflow
```

At dev time these are read via `python-dotenv` into module globals (`API_KEY_MAIN/SUB1/SUB2`) and never shown in the UI.

## Building / Shipping the Windows Executable

```batch
build_exe-win.bat
```

Builds via PyInstaller using `app-win.spec`; output goes to `C:\DIFY_build\dist\SOC_Report_Generator\` (built outside the project tree to avoid Windows MAX_PATH errors). **The `.bat` first `robocopy`s the project to a local `C:\DIFY_src` and builds from there** — when the repo lives on a OneDrive "Files On-Demand" path (e.g. a corporate `OneDrive - …\Desktop\…`), source files are online-only placeholders and PyInstaller dies with `OSError: [Errno 22] Invalid argument` while reading them; copying forces OneDrive to hydrate every file. Post-build template/`.env.example` copies also source from `C:\DIFY_src`, not the OneDrive tree. (The CI path in `build-windows.yml` runs on a clean runner, so it builds in-place without staging.) `.env` with valid keys must exist before building: `app-win.spec` reads it at build time and writes the values into `_bundled_config.py`, which PyInstaller compiles to `.pyc` and bundles — so the keys ship as bytecode, not plaintext, and end-users need no configuration. CI equivalent: `.github/workflows/build-windows.yml` (manual `workflow_dispatch`) reconstructs `.env` from GitHub Actions secrets, then runs `pyinstaller app-win.spec`.

**Frozen-vs-dev paths:** `app.py` checks `sys.frozen` and resolves `_TEMPLATE_BASE` to the exe's directory when frozen, else to the script directory. The template assets (`template_index.xlsx`, `AR_template/`, `MA_template/`) must sit alongside the exe / `app.py` accordingly — `app-win.spec` bundles them.

**Template source (`TEMPLATE_SOURCE`):** by default the MA/AR `.docx` are read from the bundled `_BUNDLED_{AR,MA}_DIR` next to the exe. `_resolve_template_dirs()` (called once right after `st.set_page_config`, which sets the `AR_TEMPLATE_DIR`/`MA_TEMPLATE_DIR` globals `resolve_template` reads) supports three alternatives: `feishu` and `sharepoint` download the latest templates from a Feishu (Lark) Drive folder / the EY SharePoint library at startup, and `onedrive` reads from a locally synced/mapped copy at `TEMPLATE_BASE_PATH`. **`feishu` mode uses the app-credential model** (the same shape as the Dify keys): `_feishu_token()` exchanges `FEISHU_APP_ID`/`FEISHU_APP_SECRET` for a `tenant_access_token`, then `_sync_feishu_templates()` lists (`/drive/v1/files?folder_token=…`, type=="file") + downloads (`/drive/v1/files/{token}/download`) the `.docx` from `FEISHU_{MA,AR}_FOLDER_TOKEN` — the folders must be shared with the app as a reader. Files must be **raw uploads** (type `"file"`), not native Feishu docs (type `"docx"`/`"doc"`), which the API can't download byte-for-byte. **Feishu can also source the letterheads and `template_index.xlsx`** via optional `FEISHU_LETTERHEAD_FOLDER_TOKEN` / `FEISHU_TEMPLATE_INDEX_FOLDER_TOKEN`: `_sync_feishu_templates()` returns a 6-tuple adding `letterhead_dir|None` + `index_path|None`, and the module-level resolver overrides `LETTERHEAD_DIR` / `TEMPLATE_INDEX` when present. These extras are **best-effort** — if their token is unset or the fetch fails, that asset stays bundled (noted in the warning) while MA/AR still load. Shared helpers: `_feishu_list_files(token, folder_token, exts)` (returns `(name, token, modified_time)` per raw file) and `_feishu_fetch_into(token, folder_token, dest, exts)` — **incremental + concurrent**: a `.manifest.json` in `dest` keyed on token+`modified_time` skips unchanged files, deletes locally any the source dropped (deletions still propagate), and downloads what's left via a `ThreadPoolExecutor` (≤8 workers; the manifest is written only after all downloads succeed, so failures retry next run). **SharePoint mode uses no stored credential / app registration** — `_sp_session()` reuses the user's existing browser sign-in via `browser_cookie3` (optional dep), `_sync_sharepoint_templates()` lists + downloads the `.docx` via the SharePoint REST API. Both online syncs are memoised with `@st.cache_resource` (sidebar "Refresh templates" button calls `.clear()`), cache into a temp dir, reject any non-`PK`/`.docx` payload, and on any failure (auth, network, missing/unshared folder) fall back to bundled templates with a UI warning (`TEMPLATE_DIR_WARNING`). `template_index.xlsx` and `LETTERHEAD_DIR` stay bundled except in `feishu` mode when their folder tokens are set (see above); SharePoint/onedrive modes keep them bundled. Per-machine settings (`TEMPLATE_SOURCE`, `TEMPLATE_BASE_PATH`, `SHAREPOINT_*`, `FEISHU_*`) come from a runtime `.env` next to the exe — `launcher-win.py` loads it with `override=False` so the baked API keys win — and are **not** compiled into `_bundled_config.py` (note `FEISHU_APP_SECRET` *is* a secret and should be baked like the Dify keys before shipping). The SharePoint path is untested against the live tenant; verify on a real machine.

## Architecture

Everything lives in one ~2700-line file: `app.py`. There is no backend server of our own — the app calls the Dify HTTP API directly. It has two largely separable concerns:

1. **Dify orchestration** — calling the three workflows and threading their outputs.
2. **`.docx` rendering** — turning the AI output into a Word document, optionally merged into pre-authored EY templates. This is the larger and more intricate half.

### Dify Orchestration

`run_workflow(inputs, api_base, api_key, status_placeholder=None)` POSTs to `/workflows/run` in **streaming (SSE)** mode with `timeout=(30, 1800)` (30s connect / 1800s read) and `verify=False` (intentional, internal use). It parses SSE events itself, updates the UI on `node_finished`, raises on `node_finished status=failed` / `workflow_finished status=failed` / `error` events, and returns the `outputs` dict from `workflow_finished`. It deliberately drops large event payloads before any Streamlit call (the MAIN end-node carries all 20 fields, hundreds of KB).

`upload_file(...)` multipart-POSTs to `/files/upload` (also `verify=False`) and returns `upload_file_id`. `to_str(v)` coerces any workflow output to `str` (`None`→`""`, list/dict→JSON with `ensure_ascii=False`) before it is passed to the next step.

**Three-step data flow** (all three run from a single **"Run All Steps (1 → 2 → 3)"** button — not three separate clicks):

1. Upload files + fill form → MAIN → `main_outputs` (20 fields)
2. SUB1 reads `main_outputs` → `sub1_outputs`
3. SUB2 reads both (`inputs_sub2`) → `final_result` (markdown string)

The MAIN→SUB1→SUB2 20-field contract is the key coupling with `DIFY-backend`; see `../CLAUDE.md`. **SUB2 fallback pattern:** SUB2 inputs use `to_str(sub1_outputs.get(key) or user_inputs.get(key))` so fields SUB1 didn't modify still flow through. `UER_json` and `cuec_preformatted` are read straight from `main_outputs` in `inputs_sub2` (they bypass SUB1). **Error handling:** Dify may return graceful error fields rather than HTTP errors — `outputs.get("Error")` / `outputs.get("Error_ORG")` are checked after each step and surfaced before stopping.

### `.docx` Rendering — Two Modes

Controlled by the **"Generate complete report (MA + AR + main sections)"** checkbox (`generate_complete`):

- **Plain mode** (unchecked): `markdown_to_docx(md_text, language)` renders SUB2's `final_result` directly to `bytes` — headings H1–H4, bullet/numbered lists, Markdown tables (incl. a `|`-delimited "numpipe" variant), bold/italic via `_inline()` (regex `**bold**`/`*italic*`; no nested/escaped asterisks). Dual fonts are applied to every run: Times New Roman (Latin) + 黑体 (CJK) via `_apply_fonts` / `_set_style_fonts`.

- **Complete mode** (checked, default): the final document is `merge_docx_sections(ma_bytes, ar_bytes, dify_bytes)` then `enforce_line_spacing(...)`, where `ma_bytes`/`ar_bytes` come from filling **EY templates** and `dify_bytes` is the plain-mode render. There is also a **"Generate MA + AR only"** button producing just `merge_docx_sections(ma, ar)` (stored under session key `ma_ar_only`).

#### EY template pipeline (complete mode)

This is the non-obvious core. Pre-authored EY Word templates live in `AR_template/` (Auditor's Report) and `MA_template/` (Management Assertion). `template_index.xlsx` is a lookup spreadsheet with `AR` and `MA` sheets; `resolve_template(report_type, standard, sso, language, sheet)` maps UI selections (report type, attestation standard, SSO strategy, EN/CN) to a **WP number**, then finds the `.docx` in the template dir whose filename starts with that number. It returns `(wp_no, path)` or `(None, error_message_string)` so the UI can show match/error status live as the form changes.

`fill_and_process_template(template_path, subs, flags, language)` opens the template, performs placeholder substitution (`build_substitutions(ui, tc)` builds the EN+CN placeholder→value dict; `build_flags(tc)` builds boolean deletion flags from the complete-report settings), and conditionally removes template paragraphs guarded by `[or ...]` / `（注：...）` markers. The remaining helpers operate on **raw WordprocessingML** (the docx unzipped to XML): `_clean_docx_bytes` / `_apply_xml_cleaning` / `_reject_format_changes` strip tracked-change and `pPrChange`/`rPrChange` format-change records; `_remap_extra_numbering` / `_inject_numbering` keep list `numId`s from colliding when documents are concatenated; `enforce_line_spacing` normalizes spacing. A recurring hazard the code works around: EY's `Normal` style is **bold**, so Dify-generated sections merged into an EY document inherit bold unless explicitly reset — several helpers exist solely to defend against this and other style/numbering bleed across the merge. Treat this cluster of functions as one tightly coupled unit.

### Session State

The sidebar **Reset** clears all of: `main_outputs`, `sub1_outputs`, `final_result`, `user_inputs`, `template_config`, `ma_ar_only`, `final_bytes`, `final_filename`. `template_config` stores the complete-report settings (standard, signing date/city, addressee, CUEC/SSO-CC choices, resolved template paths, and the boolean toggles). The final document is cached under `final_bytes`/`final_filename` so it isn't rebuilt on every Streamlit rerun; download filename is `{Co_short_name}_{Report_type}_Report.docx` (spaces → underscores).

## Workflow Input Variables (Step 1 — MAIN)

| Key | Description |
|---|---|
| `File_input` | List of `{transfer_method, upload_file_id, type}` dicts |
| `Report_type` | `"SOC1 TYPE1"`, `"SOC1 TYPE2"`, `"SOC2 TYPE1"`, `"SOC2 TYPE2"` |
| `Output_language` | `"English"`, `"中文"`, `"Both"` |
| `is_Security`, `is_Availability`, `is_Processing_Integrity`, `is_Confidentiality`, `is_Privacy` | Boolean TSC flags (SOC2 only; sent for all report types but only used by Dify when applicable) |
| `is_CUEC`, `is_UER` | Booleans — include the Complementary User Entity Controls / User Entity Responsibilities sections. Checkbox defaults: CUEC on for SOC1, UER on for SOC2. Declared on MAIN but actually consumed by SUB2 (the frontend re-sends them in `inputs_sub2`). |
| `Company_name`, `Co_short_name`, `Industry`, `System_or_service_name`, … | Company/service metadata fields |

The complete-report (MA/AR) settings — standard, signing date/city, addressee, CUEC/SSO-CC, transaction-processing/AI-scope/other-information toggles — are **not** sent to Dify; they only drive the local EY-template fill via `template_config`.
