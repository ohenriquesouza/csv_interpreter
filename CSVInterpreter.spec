# -*- mode: python ; coding: utf-8 -*-

import os

onefile = os.environ.get('CSVINTERPRETER_ONEFILE') == '1'
a = Analysis(
    ['src\\app.py'],
    pathex=['src'],
    binaries=[],
    datas=[('img/CSVinterpreter.ico', '.'), ('img/CSVinterpreter.jpg', '.')],
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
    a.binaries if onefile else [],
    a.datas if onefile else [],
    [],
    name='CSVInterpreter',
    exclude_binaries=not onefile,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['img\\CSVinterpreter.ico'],
)

if not onefile:
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        name='CSVInterpreter',
    )
