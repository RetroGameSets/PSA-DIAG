# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Ensure PySide6 dynamic libraries and data files are collected so Qt DLLs
# are bundled into the executable. This helps avoid "DLL load failed while
# importing QtCore" errors on target machines.
from PyInstaller.utils.hooks import collect_dynamic_libs, collect_data_files

_pyside6_binaries = collect_dynamic_libs('PySide6') + collect_dynamic_libs('shiboken6')
_pyside6_datas = collect_data_files('PySide6')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=_pyside6_binaries,
    datas=[
        ('updater.py', '.'),
        ('style.qss', '.'),
        ('icons/*.svg', 'icons'),
        ('icons/*.png', 'icons'),
        ('tools/*.exe', 'tools'),
        ('tools/*.dll', 'tools'),
        ('lang/*.json', 'lang'),
        ('config/*.json', 'config'),
    ] + _pyside6_datas,
    hiddenimports=[],
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
    name='PSA_DIAG',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icons/icon.ico',  # Add icon path here if you have a .ico file: icon='icons/logo.ico'
    uac_admin=True,  # Request admin privileges
)
