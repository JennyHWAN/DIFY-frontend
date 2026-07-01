# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

See also: `../CLAUDE.md` at the workspace root for how this frontend connects to the `DIFY-backend` workflows.

## Project Skills

Project-scoped skills live in `.claude/skills/` and are available to anyone working in this repo (Claude Code auto-loads them):

- **`soc-report-generator`** — authoring/editing the EY Word (`.docx`) MA/AR templates in `MA_template/` and `AR_template/` (driven by `template_index.xlsx`) and formatting the Dify-generated Section III. Covers placeholder syntax, the `[or ..]` / `【】` / `（注：…）` conditional markers, the Word-comment annotation markers (CUEC, SSO-CC, single-user-entity, AI-scope, Other-Information), and the WP-number lookup. Use when adding/changing a template, a placeholder, or a conditional rule that `app.py`'s `fill_and_process_template` must honor. Has supporting `references/` and `assets/`.
- **`frontend-design`** — guidance for distinctive, intentional visual design when building or reshaping UI (palette, typography, layout). General-purpose, not DIFY-specific.

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

Builds via PyInstaller using `app-win.spec`; output goes to `C:\DIFY_build\dist\SOC_Report_Generator\` (built outside the project tree to avoid Windows MAX_PATH errors). **The `.bat` first `robocopy`s the project to a local `C:\DIFY_src` and builds from there** — when the repo lives on a OneDrive "Files On-Demand" path (e.g. a corporate `OneDrive - …\Desktop\…`), source files are online-only placeholders and PyInstaller dies with `OSError: [Errno 22] Invalid argument` while reading them; copying forces OneDrive to hydrate every file. Post-build template/`.env.example` copies also source from `C:\DIFY_src`, not the OneDrive tree. (The CI path in `build-windows.yml` runs on a clean runner, so it builds in-place without staging.) `.env` with valid keys must exist before building: `app-win.spec` reads it at build time and writes the values into `_bundled_config.py`, which PyInstaller compiles to `.pyc` and bundles — so the keys ship as bytecode, not plaintext, and end-users need no configuration. CI equivalent: `.github/workflows/build-windows.yml` reconstructs `.env` from GitHub Actions secrets, then runs `pyinstaller app-win.spec`.

### Auto-update (Velopack)

The exe ships as a **Velopack `Setup.exe`**, not a copied folder. Velopack installs per-user to `%LocalAppData%\SOC_Report_Generator` (no admin). `launcher-win.py` runs `velopack.App().run()` first to handle install/update/uninstall hooks (required; no-op in dev). The update flow is **notify, not force**: `app.py` checks the GitHub Releases feed once per session and via a sidebar **"Check for updates"** button (`_check_for_update`; the button re-raises so the real failure surfaces instead of a misleading "up to date"), shows a notification with the new version. Clicking **"Download update in background"** starts the download+stage in a **daemon thread** (`_start_background_update`) so the app stays fully usable during the (potentially multi-hour) download; progress is shown in the sidebar, which auto-refreshes via an `st.fragment(run_every="2s")` (only while downloading) so progress + the "Restart now" prompt surface on their own without rerunning the rest of the app. When the package is staged the sidebar shows a **"Restart now to install"** button, and only that click applies + relaunches. **Nothing auto-restarts** — the download and the install are two separate, user-controlled steps. All update code is inert unless `sys.frozen` (Velopack-installed). The updater uses a **`velopack.HttpSource`** pointed at GitHub's `releases/latest/download/` redirect (`UPDATE_FEED_URL`), **not** `GithubSource`: `GithubSource` queries `api.github.com`, which on locked-down corporate networks returns 403/empty (silent "no update") even when plain release downloads work — `HttpSource` uses only `github.com`/`githubusercontent.com`. The sidebar **🔧 Update diagnostics** expander dumps `get_current_version()`/`get_app_id()`/`check_for_updates()` plus an independent `requests` probe of the feed, to tell network problems apart from Velopack-internal ones. **Corporate TLS interception:** on networks that re-sign HTTPS with an internal root CA (e.g. EY machines), the frozen app's *bundled* trust stores (certifi for `requests`, webpki for Velopack's Rust TLS) reject the cert as `UnknownIssuer` even though Windows trusts it — this is the real reason updates failed (silent under `GithubSource`, then an explicit cert error under `HttpSource`). `_install_corp_certs()` (run once at import, before any TLS call) dumps the Windows `ROOT`+`CA` stores via stdlib `ssl.enum_certificates()` to a PEM in temp and points **both** `SSL_CERT_FILE` and `REQUESTS_CA_BUNDLE` at it (`requests` and `rustls-native-certs` both consult `SSL_CERT_FILE` first), so the app trusts exactly what the OS trusts. No-op off Windows. If the machine **already** sets `SSL_CERT_FILE`/`REQUESTS_CA_BUNDLE` to a narrow single-purpose bundle (e.g. an internal platform `.cer`), that cert is **folded into** the combined bundle and the app **takes over** both vars (override, not `setdefault`) — that bundle alone lacks the interception root, so respecting it would keep updates broken (the original bug). Override is safe: the only verified HTTPS is the update feed (Dify uses `verify=False`) and the env edits are process-local; idempotent across Streamlit reruns. The diagnostics panel shows the resulting `SSL_CERT_FILE` + bundle size (a real Windows-store dump is tens of KB, not ~1–2 KB). **However, that only fixes `requests` — Velopack's bundled Rust TLS ignores `SSL_CERT_FILE` and still rejects the corporate cert as `UnknownIssuer`, so its own `check_for_updates`/`download_updates` can never work on an inspecting network.** Final design (do all networking with `requests`, let Velopack only do the local apply): `_check_for_update()` fetches `releases.win.json` with `requests`, parses the highest-version `Type=="Full"` asset (`_latest_full_asset`), and compares `_ver_tuple` against `get_current_version()` (a local, no-network Velopack call) — returning the **asset dict** (not a Velopack `UpdateInfo`). The apply is split into a slow background half and a fast foreground half. `_stage_update_background(asset)` — run in a daemon thread, tracked through the module-level `_UPDATE_STATE` dict (state `idle`/`downloading`/`staged`/`error`) since a worker thread can't touch `st.session_state` — downloads the full `.nupkg` (**deltas are disabled** — `vpk pack --delta none`; see below) + writes a `releases.win.json` into a temp staging dir via `requests`, then constructs `UpdateManager(<staging_dir>)` (a **local file source** — no TLS) and runs `check_for_updates`/`download_updates` to **assemble** the package, but **stops there**: it stashes the ready `(UpdateManager, UpdateInfo)` in the module-level `_STAGED` and flips state to `staged`. The actual install is deferred to the user's **"Restart now"** click, which sets `_do_restart` and reruns; the top of the script paints a brief full-viewport overlay and calls `_restart_to_apply()`, which writes `update_restart.flag`, calls `wait_exit_then_apply_updates(restart=True)`, then `os._exit(0)` (a direct `apply_updates_and_restart` deadlocks inside the embedded Streamlit server — the updater waits for this PID to exit while the call blocks the run thread). Velopack still verifies the package SHA against the local index. The `HttpSource` manager is now used only for the local `get_current_version` call. The build version is the root **`VERSION`** file — bundled into `_MEIPASS` (spec datas) for the sidebar label; `vpk --packVersion` is the source of truth for the actual installed version. `build_exe-win.bat` runs `vpk pack` after PyInstaller (output `C:\DIFY_build\Releases\`); `app-win.spec` `collect_all("velopack")`s the native helper lib. **Releases are cut by pushing a git tag `vX.Y.Z`** — the workflow then packs and runs `vpk upload github --publish` (needs `permissions: contents: write`; `fetch-depth: 0` so prior releases are available to build the feed). A manual `workflow_dispatch` run packages but only uploads the installer as an artifact (no published Release). **`vpk pack` runs with `--delta none` — full packages only.** Delta reconstruction (installed base + delta → full) couldn't rebuild the published full byte-for-byte here (PyInstaller builds aren't perfectly deterministic), so Velopack rejected the reassembled package with `Size did not match for …-full.partial`; this also broke *old* installed clients (which run their own version's update code, predating our delta→full fallback), so the only cure was to republish a full-only feed. Full packages are self-verifying (own SHA, no reconstruction) and the resumable `requests` downloader handles the ~95 MB fine. Existing folder-copy users must run `Setup.exe` once to migrate onto the managed install. Requires the .NET SDK for `vpk` (installed as a global dotnet tool).

**Frozen-vs-dev paths:** `app.py` checks `sys.frozen` and resolves `_TEMPLATE_BASE` to the exe's directory when frozen, else to the script directory. The template assets (`template_index.xlsx`, `AR_template/`, `MA_template/`) must sit alongside the exe / `app.py` accordingly — `app-win.spec` bundles them.

**Template source (`TEMPLATE_SOURCE`):** by default the MA/AR `.docx` are read from the bundled `_BUNDLED_{AR,MA}_DIR` next to the exe. `_resolve_template_dirs()` (called once right after `st.set_page_config`, which sets the `AR_TEMPLATE_DIR`/`MA_TEMPLATE_DIR` globals `resolve_template` reads) supports three alternatives: `feishu` and `sharepoint` download the latest templates from a Feishu (Lark) Drive folder / the EY SharePoint library at startup, and `onedrive` reads from a locally synced/mapped copy at `TEMPLATE_BASE_PATH`. **`feishu` mode uses the app-credential model** (the same shape as the Dify keys): `_feishu_token()` exchanges `FEISHU_APP_ID`/`FEISHU_APP_SECRET` for a `tenant_access_token`, then `_sync_feishu_templates()` lists (`/drive/v1/files?folder_token=…`, type=="file") + downloads (`/drive/v1/files/{token}/download`) the `.docx` from `FEISHU_{MA,AR}_FOLDER_TOKEN` — the folders must be shared with the app as a reader. Files must be **raw uploads** (type `"file"`), not native Feishu docs (type `"docx"`/`"doc"`), which the API can't download byte-for-byte. **Feishu can also source the letterheads and `template_index.xlsx`** via optional `FEISHU_LETTERHEAD_FOLDER_TOKEN` / `FEISHU_TEMPLATE_INDEX_FOLDER_TOKEN`: `_sync_feishu_templates()` returns a 6-tuple adding `letterhead_dir|None` + `index_path|None`, and the module-level resolver overrides `LETTERHEAD_DIR` / `TEMPLATE_INDEX` when present. These extras are **best-effort** — if their token is unset or the fetch fails, that asset stays bundled (noted in the warning) while MA/AR still load. Shared helpers: `_feishu_list_files(token, folder_token, exts)` (returns `(name, token, modified_time)` per raw file) and `_feishu_fetch_into(token, folder_token, dest, exts)` — **incremental + concurrent**: a `.manifest.json` in `dest` keyed on token+`modified_time` skips unchanged files, deletes locally any the source dropped (deletions still propagate), and downloads what's left via a `ThreadPoolExecutor` (≤4 workers; `_feishu_download` retries Feishu's rate-limit response — code `99991400` / HTTP 429 — with exponential backoff + jitter; the manifest is written only after all downloads succeed, so failures retry next run). **SharePoint mode uses no stored credential / app registration** — `_sp_session()` reuses the user's existing browser sign-in via `browser_cookie3` (optional dep), `_sync_sharepoint_templates()` lists + downloads the `.docx` via the SharePoint REST API. Both online syncs are memoised with `@st.cache_resource` (sidebar "Refresh templates" button calls `.clear()`), cache into a temp dir, reject any non-`PK`/`.docx` payload, and on any failure (auth, network, missing/unshared folder) fall back to bundled templates with a UI warning (`TEMPLATE_DIR_WARNING`). `template_index.xlsx` and `LETTERHEAD_DIR` stay bundled except in `feishu` mode when their folder tokens are set (see above); SharePoint/onedrive modes keep them bundled. Per-machine settings (`TEMPLATE_SOURCE`, `TEMPLATE_BASE_PATH`, `SHAREPOINT_*`, `FEISHU_*`, incl. `FEISHU_APP_SECRET`) are **baked into `_bundled_config.py` at build time** from the build-time `.env` (as a `RUNTIME_ENV` dict in `app-win.spec`), so the shipped exe behaves like dev with no plaintext `.env` beside it and the secret lives in bytecode. `launcher-win.py` applies them: it sets the baked Dify keys, loads an optional runtime `.env` next to the exe with `override=False` (so baked API keys win), then `setdefault`s the baked `RUNTIME_ENV` **after** the runtime `.env` — so a per-machine `.env` can still override the non-secret template settings, while the baked values are the working defaults. The SharePoint path is untested against the live tenant; verify on a real machine.

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
