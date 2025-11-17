# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['excel_slimmer_qt.py'],
    pathex=['.', 'backData'],
    binaries=[],
    datas=[('check_white.svg', '.')],
    hiddenimports=[
        'gui_clean_defined_names_desktop_date',
        'excel_image_slimmer_gui_v3',
        'excel_slimmer_precision_plus',
    ],
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
    name='ExcelSlimmer',
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
    icon=['ExcelSlimmer.ico'],
)
