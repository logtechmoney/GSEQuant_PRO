# -*- mode: python ; coding: utf-8 -*-
# ─────────────────────────────────────────────────────────────────────────────
# launcher.spec — PyInstaller spec para GSEQuant Launcher
# El launcher resultante pesa ~5-8 MB (sin PyQt5, solo tkinter + stdlib)
#
# Compilar con:
#     pyinstaller launcher.spec --noconfirm
# ─────────────────────────────────────────────────────────────────────────────

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('gse_app_icon.png', '.'),
        ('gta.png', '.'),
    ],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        '_tkinter',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Excluimos todo lo que NO necesita el launcher para mantenerlo liviano
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'numpy',
        'pandas',
        'matplotlib',
        'scipy',
        'networkx',
        'PIL',
        'cv2',
        'sklearn',
    ],
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
    name='GSEQuant_Launcher',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # Sin ventana de consola
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='gse_app_icon.png',
)
