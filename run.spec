# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

block_cipher = None

BASE_DIR = Path(__file__).parent.resolve()

datas = [
    (BASE_DIR / 'assets' / 'logo.png', 'assets'),
    (BASE_DIR / 'data' / 'zip_lookup.csv', 'data'),
    (BASE_DIR / 'data' / 'master_client_list.xlsx', 'data'),
]

for signature in (BASE_DIR / 'assets' / 'signatures').glob('*.png'):
    datas.append((signature, 'assets/signatures'))
a = Analysis(
    ['run.py'],
    pathex=[str(BASE_DIR)],
    binaries=[],
    datas=datas,
    hiddenimports=['AutoMailerPro', 'pandas', 'docx', 'fuzzywuzzy', 'Levenshtein', 'tkinter', 'ttkthemes'],
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
    name='AutoMailerPro',
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
    icon='logo.ico',  # Remove if logo.ico doesn't exist
)