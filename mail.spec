# -*- mode: python ; coding: utf-8 -*-


block_cipher = None
from pathlib import Path

p = Path('.env')
if not p.exists():
    raise FileNotFoundError('找不到檔案：.env')
for line in p.read_text(encoding='utf-8').split('\n'):
    if '=' not in line:
        continue
    line = line.split('=')
    key = line[0]
    value = '='.join(line[1:])
    globals()[key] = value
if not Path('dist').exists():
    Path('dist').mkdir()


a = Analysis(
    ['mail.py'],
    pathex=[Path(SPECPATH).absolute()],
    binaries=[],
    datas=[('.env', '.')],
    hiddenimports=['pkg_resources.py2_warn'],
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
    [],
    exclude_binaries=True,
    name='mail',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=['vcruntime140.dll'],
    name=f'批量郵件發送機器人V{VERSION}',
)
