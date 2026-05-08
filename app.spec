# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec-файл для сборки Конвертер_Jira_CSV_XLSX.exe

Сборка:
    python -m PyInstaller app.spec

Результат:  dist/Конвертер_Jira_CSV_XLSX.exe  (~15–20 МБ, одиночный файл)
"""

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/icon.ico',     'assets'),
        ('assets/icon_256.png', 'assets'),
        ('assets/icon_128.png', 'assets'),
        ('assets/icon_64.png',  'assets'),
        ('assets/icon_48.png',  'assets'),
        ('assets/icon_32.png',  'assets'),
        ('assets/icon_16.png',  'assets'),
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.worksheet.table',
        'tkinterdnd2',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'numpy.distutils',
        'IPython',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Конвертер_Jira_CSV_XLSX',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,           # без консольного окна
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    icon='assets/icon.ico',  # иконка Core-Cat
)
