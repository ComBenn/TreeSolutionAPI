# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['.venv\\files\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('.venv\\files\\keywords_technische_accounts.txt', '.'), ('.venv\\files\\batch_export_tracker.json', '.'), ('.venv\\files\\ui_state.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='TreeSolutionAPI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
