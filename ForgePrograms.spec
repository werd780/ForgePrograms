# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

# Collect tkinter submodules (helps with filedialog/messagebox issues in frozen builds)
tk_hidden = collect_submodules("tkinter")
openpyxl_hidden = collect_submodules("openpyxl")
fillpdf_hidden = collect_submodules("fillpdf")
pdfrw_hidden = collect_submodules("pdfrw")

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("Genner1150", "Genner1150"),
        ("Inventory", "Inventory"),
    ],
    hiddenimports=tk_hidden + openpyxl_hidden + fillpdf_hidden + pdfrw_hidden + [
        "tkinter.filedialog",
        "tkinter.messagebox",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="ForgePrograms",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,   # Tkinter GUI app; no console window
    disable_windowed_traceback=False,
)

# onefile=True makes a single executable in dist/
# If you ever want onedir builds for debugging, set onefile=False
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="ForgePrograms",
)