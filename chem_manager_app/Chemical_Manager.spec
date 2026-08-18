# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
import os

selenium_datas, selenium_binaries, selenium_hiddenimports = collect_all('seleniumbase')
guide_path = os.path.abspath(os.path.join(SPECPATH, '..', 'program_guide.md'))
icon_ico_path = os.path.abspath(os.path.join(SPECPATH, '..', 'chemical-reagent-manager-icon.ico'))
icon_png_path = os.path.abspath(os.path.join(SPECPATH, '..', 'chemical-reagent-manager-icon.png'))


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=selenium_binaries,
    datas=selenium_datas + [(guide_path, '.'), (icon_png_path, '.')],
    hiddenimports=selenium_hiddenimports + ['win32timezone', 'pythoncom', 'pywintypes', 'pymupdf',
                   'seleniumbase', 'scrapers.aldrich', 'scrapers.tci',
                   'scrapers.thermofisher', 'scrapers.abcam', 'scrapers.coa_downloader'],
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
    name='Chemical_Manager',
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
    icon=icon_ico_path,
)
