# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

datas = [('gta.png', '.'), ('gse_app_icon.png', '.'), ('config_defaults', 'config_defaults'), ('Input', 'Input')]
binaries = []
hiddenimports = []
hiddenimports += collect_submodules('PyQt5')
tmp_ret = collect_all('PyQt5')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# pandas + openpyxl: se importan condicionalmente en el código (try/except),
# por lo que PyInstaller NO los detecta automáticamente. Se fuerzan aquí.
hiddenimports += collect_submodules('pandas')
tmp_pandas = collect_all('pandas')
datas += tmp_pandas[0]; binaries += tmp_pandas[1]; hiddenimports += tmp_pandas[2]

hiddenimports += collect_submodules('openpyxl')
tmp_openpyxl = collect_all('openpyxl')
datas += tmp_openpyxl[0]; binaries += tmp_openpyxl[1]; hiddenimports += tmp_openpyxl[2]

# Excluded packages that we know GSEQuant does not use and take up a lot of space
excludes = [
    'tkinter', '_tkinter', 'matplotlib', 'scipy', 'sklearn',
    'jupyter', 'notebook', 'IPython', 'pydoc', 'xmlrpc', 'http.server',
    'PyQt6', 'PySide6', 'PySide2', 'jedi', 'setuptools'
]

a = Analysis(
    ['GSEQuant_int_fixed.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=2,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='GSEQuant_PRO',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,     
    upx=False,       
    console=False,  # Oculta la pantallita negra de cmd
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['gsequant.ico'],
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='GSEQuant_PRO',
)
