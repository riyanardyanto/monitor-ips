# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files


flet_datas = collect_data_files('flet')
flet_desktop_datas = collect_data_files('flet_desktop')
project_datas = [
    ('field_cell.txt', '.'),
    ('assets/app_icon.ico', 'assets'),
    ('assets/ips.html', 'assets'),
    ('docs/ips-checker-user-guide.pdf', 'docs'),
    ('docs/ips-generator-user-guide.pdf', 'docs'),
]

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=flet_datas + flet_desktop_datas + project_datas,
    hiddenimports=[],
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
    name='IPS Checker',
    icon='assets/app_icon.ico',
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
)
