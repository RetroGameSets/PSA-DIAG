# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Ensure PySide6 dynamic libraries and data files are collected so Qt DLLs
# are bundled into the executable. This helps avoid "DLL load failed while
# importing QtCore" errors on target machines.

# Temporarily disable collecting PySide6 data for test builds
# (this reduces bundle size and avoids PySide-related packaging during tests)
_pyside6_binaries = []
_pyside6_datas = []
# PySide6 collection intentionally disabled for test builds

# Prepare datas list
datas_list = [
    ('style.qss', '.'),
    ('icons/*.svg', 'icons'),
    ('icons/*.png', 'icons'),
    ('icons/*.ico', 'icons'),
    ('tools/*.exe', 'tools'),
    ('tools/*.dll', 'tools'),
    ('lang/*.json', 'lang'),
    ('config/*.json', 'config'),
    ('config.py', '.'),
]

datas = datas_list + _pyside6_datas

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=_pyside6_binaries,
    datas=datas,
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
