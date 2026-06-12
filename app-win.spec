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

# ── Include app source files ───────────────────────────────────────────────────
datas += [("app.py", ".")]

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

with open("_bundled_config.py", "w") as _f:
    _f.write(
        f"API_BASE_URL = {repr(_env.get('DIFY_API_BASE_URL', 'https://api.dify.ai/v1'))}\n"
        f"API_KEY_MAIN = {repr(_env.get('DIFY_API_KEY_MAIN', ''))}\n"
        f"API_KEY_SUB1 = {repr(_env.get('DIFY_API_KEY_SUB1', ''))}\n"
        f"API_KEY_SUB2 = {repr(_env.get('DIFY_API_KEY_SUB2', ''))}\n"
    )
datas += [("_bundled_config.py", ".")]

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
