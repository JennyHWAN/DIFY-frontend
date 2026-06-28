# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for SOC Report Generator.
Bundles the Streamlit app + all dependencies into a self-contained
Windows folder (dist/SOC_Report_Generator/).

Build command (run build_exe.bat or):
    pyinstaller app.spec
"""

from PyInstaller.utils.hooks import collect_data_files, collect_all
import os

datas = []
binaries = []
hiddenimports = []

# ── Collect all Streamlit runtime files (static assets, templates, etc.) ──────
tmp = collect_all("streamlit")
datas    += tmp[0]
binaries += tmp[1]
hiddenimports += tmp[2]

# ── Other packages that need their data files ──────────────────────────────────
for pkg in ("docx", "altair", "pyarrow", "pydeck"):
    try:
        datas += collect_data_files(pkg)
    except Exception:
        pass

# browser_cookie3 (optional; only used for TEMPLATE_SOURCE=sharepoint). collect_all
# pulls in its crypto/lz4 backends so the frozen exe can read browser cookies. Guarded
# so the build still works if the package isn't installed — the app falls back to
# bundled templates when it's missing.
try:
    _bc = collect_all("browser_cookie3")
    datas += _bc[0]
    binaries += _bc[1]
    hiddenimports += _bc[2]
except Exception:
    pass

# velopack (Windows auto-update) ships a native helper library that must be
# bundled alongside the Python bindings, so collect_all pulls its data/binaries.
try:
    _vp = collect_all("velopack")
    datas += _vp[0]
    binaries += _vp[1]
    hiddenimports += _vp[2]
except Exception:
    pass

# ── Include app source files ───────────────────────────────────────────────────
datas += [("app.py", ".")]

# Bundle the VERSION file so the app can display the current build version in the
# sidebar (read from _MEIPASS at runtime). vpk's packVersion is the source of truth
# for the actual installed version; this is just for the UI label.
if os.path.exists("VERSION"):
    datas += [("VERSION", ".")]

# ── Embed API keys as compiled bytecode (not a readable plaintext file) ────────
# Read .env at build time, write values into a Python module, then bundle it.
# PyInstaller compiles _bundled_config.py to .pyc — keys are in bytecode, not
# a plaintext file that users can open from the _internal folder.
try:
    from dotenv import dotenv_values as _dv
    _env = _dv(".env")
except Exception:
    raise FileNotFoundError(".env not found — create it with your API keys before building.")

for _key in ("DIFY_API_KEY_MAIN", "DIFY_API_KEY_SUB1", "DIFY_API_KEY_SUB2"):
    if not _env.get(_key):
        raise ValueError(f"{_key} is missing or empty in .env")

# Per-machine / template settings (incl. the Feishu app secret) are baked too, so
# the shipped exe behaves like dev without a plaintext .env beside it — and the
# secret lives in bytecode, not a readable file. Only keys actually present in .env
# are written; a runtime .env can still override the non-secret ones (the launcher
# loads it before applying these as defaults). API keys stay baked-wins.
_OPTIONAL = (
    "TEMPLATE_SOURCE", "TEMPLATE_BASE_PATH", "MA_TEMPLATE_SUBPATH", "AR_TEMPLATE_SUBPATH",
    "FEISHU_API_BASE", "FEISHU_APP_ID", "FEISHU_APP_SECRET",
    "FEISHU_MA_FOLDER_TOKEN", "FEISHU_AR_FOLDER_TOKEN",
    "FEISHU_LETTERHEAD_FOLDER_TOKEN", "FEISHU_TEMPLATE_INDEX_FOLDER_TOKEN",
    "SHAREPOINT_SITE_URL", "SHAREPOINT_MA_FOLDER", "SHAREPOINT_AR_FOLDER",
)
_runtime_env = {k: _env[k] for k in _OPTIONAL if _env.get(k)}

with open("_bundled_config.py", "w") as _f:
    _f.write(
        f"API_BASE_URL = {repr(_env.get('DIFY_API_BASE_URL', 'https://api.dify.ai/v1'))}\n"
        f"API_KEY_MAIN = {repr(_env.get('DIFY_API_KEY_MAIN', ''))}\n"
        f"API_KEY_SUB1 = {repr(_env.get('DIFY_API_KEY_SUB1', ''))}\n"
        f"API_KEY_SUB2 = {repr(_env.get('DIFY_API_KEY_SUB2', ''))}\n"
        f"RUNTIME_ENV = {repr(_runtime_env)}\n"
    )
print(f"[spec] baked template settings: {sorted(_runtime_env)}")
datas += [("_bundled_config.py", ".")]

# ── Generate the protobuf runtime hook at build time ───────────────────────────
# PROTOCOL_BUFFERS_PYTHON_IMPLEMENTATION must be set before protobuf is imported
# in the frozen app. Writing the hook here (instead of relying on a committed
# rthook_protobuf.py being present) makes the build self-healing: it never fails
# with "FileNotFoundError: rthook_protobuf.py" when the source folder is a
# OneDrive copy that hasn't synced that file down yet.
with open("rthook_protobuf.py", "w") as _f:
    _f.write('import os\n')
    _f.write('os.environ["PROTOCOL_BUFFERS_PYTHON_IMPLEMENTATION"] = "python"\n')

# ── Analysis ───────────────────────────────────────────────────────────────────
a = Analysis(
    ["launcher-win.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + [
        # Streamlit internals
        "streamlit.web.cli",
        "streamlit.runtime.scriptrunner.magic_funcs",
        "streamlit.components.v1",
        # App dependencies
        "docx",
        "docx.oxml.ns",
        "docx.oxml",
        "requests",
        "dotenv",
        "re",
        "io",
        "zipfile",
        "xml.etree.ElementTree",
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.utils",
        "copy",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=["rthook_protobuf.py"],
    excludes=["matplotlib", "scipy", "PIL", "cv2", "tensorflow", "torch"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,          # use COLLECT (onedir) — faster startup
    name="SOC_Report_Generator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,                   # keep console window so errors are visible
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="SOC_Report_Generator",    # output folder: dist/SOC_Report_Generator/
)
