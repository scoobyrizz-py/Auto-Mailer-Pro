# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path

block_cipher = None

spec_path = Path(__file__) if "__file__" in globals() else Path(sys.argv[0])
BASE_DIR = spec_path.parent.resolve()

datas = [
    (str(BASE_DIR / 'assets' / 'logo.png'), 'assets'),
    (str(BASE_DIR / 'assets' / 'logo.ico'), 'assets'),
    (str(BASE_DIR / 'data' / 'zip_lookup.csv'), 'data'),
    (str(BASE_DIR / 'data' / 'master_client_list.xlsx'), 'data'),
]

for signature in (BASE_DIR / 'assets' / 'signatures').glob('*.png'):
    datas.append((str(signature), 'assets/signatures'))
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
    icon=str(BASE_DIR / 'assets' / 'logo.ico'),
)
