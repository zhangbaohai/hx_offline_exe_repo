# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files
app_name = "华夏离线批量编辑器"

hidden = []
hidden += collect_submodules("pandas")
hidden += collect_submodules("openpyxl")

datas = []
datas += collect_data_files("openpyxl")
datas += [("README.txt", "."), ("requirements.txt", ".")]

a = Analysis(["app_exact.py"], pathex=[], binaries=[], datas=datas, hiddenimports=hidden, hookspath=[], hooksconfig={}, runtime_hooks=[], excludes=[], win_no_prefer_redirects=False, win_private_assemblies=False, cipher=None, noarchive=False)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)
exe = EXE(pyz, a.scripts, [], exclude_binaries=True, name=app_name, debug=False, bootloader_ignore_signals=False, strip=False, upx=True, console=False, icon="icon.ico")
coll = COLLECT(exe, a.binaries, a.zipfiles, a.datas, strip=False, upx=True, upx_exclude=[], name=app_name)