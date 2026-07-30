# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files
import sys

block_cipher = None

# ------------------------------------------------------------
# Analysis
# ------------------------------------------------------------
# 1. Define base assets that every OS needs
app_datas = [
    ('TreeX4GD-template.html', '.'),
    ('TreeX4GD-help.html', '.'),
    ('TreeX4GD-help-images', 'TreeX4GD-help-images'),
]

# 2. Conditionally append OS-specific icons
if sys.platform.startswith('win'):
    app_datas.append(('TreeX4GD.ico', '.'))
elif sys.platform.startswith('linux'):
    app_datas.append(('TreeX4GD.png', '.'))

# 3. Prepare Analysis
a = Analysis(
    ['TreeX4GD.py'],
    pathex=[],
    binaries=[],
    datas=app_datas 
    + collect_data_files('googleapiclient')
    + collect_data_files('google.auth'),
    # NOTE: google_auth_httplib2 is usually code-only and doesn't require data file collection.
    # If you encounter "Resource Not Found" errors related to Google authentication, 
    # delete the comma above and uncomment the line below:
    # + collect_data_files('google_auth_httplib2'), 
    hiddenimports=[
        'requests',
        'chardet',
        'charset_normalizer',
        # WARNING: The following hidden import is a workaround for specific versions 
        # of requests/charset_normalizer. Remove if it causes "ModuleNotFound" during build.
        'charset_normalizer.md__mypyc',
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

# ------------------------------------------------------------
# PYZ
# ------------------------------------------------------------
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ------------------------------------------------------------
# EXE (All platforms)
# ------------------------------------------------------------
exe = EXE(
    pyz,
    a.scripts,
    # For Windows/Linux single-file, we must include these here:
    a.binaries if not sys.platform.startswith('darwin') else [],
    a.zipfiles if not sys.platform.startswith('darwin') else [],
    a.datas if not sys.platform.startswith('darwin') else [],
    
    exclude_binaries=True if sys.platform.startswith('darwin') else False,
    name='TreeX4GD',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX is only enabled for Windows to avoid stability/signing issues on Linux/macOS
    upx=True if sys.platform.startswith('win') else False,
    upx_exclude=(
        ['vcruntime*.dll', 'python*.dll']
        if sys.platform.startswith('win')
        else []
    ),
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    icon='TreeX4GD.ico' if sys.platform.startswith('win') else None,
    # macOS uses .icns in BUNDLE below and icons are not well supported in Linux
)

# ------------------------------------------------------------
# macOS-only app bundle
# ------------------------------------------------------------
if sys.platform.startswith('darwin'):
    app = BUNDLE(
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        name='TreeX4GD.app',
        icon='TreeX4GD.icns',
        bundle_identifier='com.github.luismica.TreeX4GD',
    )