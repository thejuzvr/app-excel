# -*- mode: python ; coding: utf-8 -*-
import os
import shutil
from PyInstaller.utils.hooks import collect_all

datas = [('templates', 'templates'), ('assets', 'assets')]
binaries = []
hiddenimports = ['babel.numbers']

for pkg in ['customtkinter', 'tkcalendar']:
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    [],
    exclude_binaries=True,
    name='Yahoo',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/ico.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Yahoo',
)

# Копирование папок templates и assets напрямую рядом с .exe
spec_dir = os.path.dirname(os.path.abspath(SPEC)) if 'SPEC' in globals() and SPEC else os.path.abspath('.')
dist_target_dir = os.path.join(spec_dir, 'dist', 'Yahoo')

for folder in ['templates', 'assets']:
    src_path = os.path.join(spec_dir, folder)
    dst_path = os.path.join(dist_target_dir, folder)
    if os.path.exists(src_path):
        try:
            if os.path.exists(dst_path):
                shutil.rmtree(dst_path)
            shutil.copytree(src_path, dst_path)
        except Exception as e:
            print(f"Warning: could not copy {folder} to dist: {e}")
