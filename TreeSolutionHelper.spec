# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src\\treesolution_helper\\files\\main.py'],
    pathex=['src\\treesolution_helper\\files'],
    binaries=[],
    datas=[
        ('README.md', '.'),
        ('src\\treesolution_helper\\files\\keywords_technische_accounts.txt', '.'),
    ],
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
    name='TreeSolutionHelper',
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
