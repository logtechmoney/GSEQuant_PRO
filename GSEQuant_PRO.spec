# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

datas = [('gta.png', '.'), ('gse_app_icon.png', '.'), ('config_defaults', 'config_defaults'), ('Input', 'Input')]
binaries = []
hiddenimports = []
hiddenimports += collect_submodules('PyQt5')
tmp_ret = collect_all('PyQt5')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]

# pandas + openpyxl: importados condicionalmente (try/except), PyInstaller NO los detecta solo
hiddenimports += collect_submodules('pandas')
tmp_pandas = collect_all('pandas')
datas += tmp_pandas[0]; binaries += tmp_pandas[1]; hiddenimports += tmp_pandas[2]

hiddenimports += collect_submodules('openpyxl')
tmp_openpyxl = collect_all('openpyxl')
datas += tmp_openpyxl[0]; binaries += tmp_openpyxl[1]; hiddenimports += tmp_openpyxl[2]

# networkx: importado global pero tiene submódulos lazy que PyInstaller pierde
hiddenimports += collect_submodules('networkx')
tmp_nx = collect_all('networkx')
datas += tmp_nx[0]; binaries += tmp_nx[1]; hiddenimports += tmp_nx[2]

# requests: importado condicionalmente (try/except)
hiddenimports += collect_submodules('requests')
tmp_req = collect_all('requests')
datas += tmp_req[0]; binaries += tmp_req[1]; hiddenimports += tmp_req[2]

# Dependencias de requests
for _pkg in ('urllib3', 'certifi', 'charset_normalizer', 'idna'):
    try:
        hiddenimports += collect_submodules(_pkg)
        _tmp = collect_all(_pkg)
        datas += _tmp[0]; binaries += _tmp[1]; hiddenimports += _tmp[2]
    except Exception:
        pass

# numpy: dependencia core de pandas
hiddenimports += collect_submodules('numpy')
tmp_np = collect_all('numpy')
datas += tmp_np[0]; binaries += tmp_np[1]; hiddenimports += tmp_np[2]

# Paquetes excluidos: no usados + test suites que inflan el build
excludes = [
    'tkinter', '_tkinter', 'matplotlib', 'scipy', 'sklearn',
    'jupyter', 'notebook', 'IPython', 'pydoc', 'xmlrpc', 'http.server',
    'PyQt6', 'PySide6', 'PySide2', 'jedi', 'setuptools',
    'pandas.tests',
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
    console=False,
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
