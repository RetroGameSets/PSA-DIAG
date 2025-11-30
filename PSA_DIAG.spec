# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Ensure PySide6 dynamic libraries and data files are collected so Qt DLLs
# are bundled into the executable. This helps avoid "DLL load failed while
# importing QtCore" errors on target machines.
import os
try:
    from PyInstaller.utils.hooks import collect_data_files, Tree
except Exception:
    from PyInstaller.utils.hooks import collect_data_files
    Tree = None

# Temporarily disable collecting PySide6 data for test builds
# (this reduces bundle size and avoids PySide-related packaging during tests)
_pyside6_binaries = []
_pyside6_datas = []
# PySide6 collection intentionally disabled for test builds

# Prepare datas list and include updater onedir distribution whether Tree is available or not
datas_list = [
    ('updater.py', '.'),
    ('style.qss', '.'),
    ('icons/*.svg', 'icons'),
    ('icons/*.png', 'icons'),
    ('tools/*.exe', 'tools'),
    ('tools/*.dll', 'tools'),
    ('lang/*.json', 'lang'),
    ('config/*.json', 'config'),
]

# If PyInstaller provides Tree, use it to include the whole folder. Otherwise fall back to manual listing.
updater_dir = os.path.join('tools', 'updater')
if Tree is not None and os.path.isdir(updater_dir):
    datas_list.append(Tree(updater_dir, prefix='tools/updater'))
elif os.path.isdir(updater_dir):
    # Walk the updater directory and add each file
    for root, _, files in os.walk(updater_dir):
        rel_root = os.path.relpath(root, updater_dir)
        for fn in files:
            src = os.path.join(root, fn)
            # destination inside the bundle should preserve the updater folder structure
            dest_dir = os.path.join('tools', 'updater', rel_root) if rel_root != '.' else os.path.join('tools', 'updater')
            datas_list.append((src, dest_dir))

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
