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
if os.path.exists(".env.example"):
    datas += [(".env.example", ".")]

# ── Analysis ───────────────────────────────────────────────────────────────────
a = Analysis(
    ["launcher.py"],
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
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
