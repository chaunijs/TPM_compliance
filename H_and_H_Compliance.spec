# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = [('C:\\ProgramData\\miniforge3\\Library\\bin\\ffi-8.dll', '.'), ('C:\\ProgramData\\miniforge3\\Library\\bin\\libcrypto-3-x64.dll', '.'), ('C:\\ProgramData\\miniforge3\\Library\\bin\\libssl-3-x64.dll', '.'), ('C:\\ProgramData\\miniforge3\\Library\\bin\\liblzma.dll', '.'), ('C:\\ProgramData\\miniforge3\\Library\\bin\\libbz2.dll', '.'), ('C:\\ProgramData\\miniforge3\\Library\\bin\\sqlite3.dll', '.')]
hiddenimports = ['pyxlsb', 'pandas', 'pandas._libs.tslibs.base']
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('xlsxwriter')
hiddenimports += collect_submodules('rich')
hiddenimports += collect_submodules('questionary')
hiddenimports += collect_submodules('prompt_toolkit')
tmp_ret = collect_all('polars')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('polars_runtime_32')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('fastexcel')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['C:\\Users\\Teerapat.Haeranyikan\\TPM-build\\TPM_compliance\\tpm_compliance_h_and_h.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['matplotlib', 'tkinter', 'PyQt5', 'PyQt6', 'IPython', 'jupyter', 'pytest', 'hypothesis', 'pandas.tests', 'pandas.io.tests', 'numpy.tests', 'asyncssh'],
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
    name='H_and_H_Compliance',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
