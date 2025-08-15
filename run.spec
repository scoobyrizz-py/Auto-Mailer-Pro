# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['run.py'],
    pathex=['C:\\Users\\scoob\\OneDrive\\Desktop\\Auto Mailer Pro'],
    binaries=[],
    datas=[
        ('logo.png', '.'),
        ('signature_brian.png', '.'),
        ('signature_bob.png', '.'),
        ('signature_kyle.png', '.'),
        ('signature_julie.png', '.'),
        ('zip_lookup.csv', '.'),
    ],
    hiddenimports=['AutoMailerPro_v5_1', 'pandas', 'docx', 'fuzzywuzzy', 'Levenshtein', 'tkinter', 'ttkthemes'],
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