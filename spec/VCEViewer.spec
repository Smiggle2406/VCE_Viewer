# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['..\\vce_viewer.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\Bailey\\PycharmProjects\\Vce_Viewer\\.venv\\Lib\\site-packages\\PyQt6\\Qt6\\plugins', 'PyQt6/Qt6/plugins'), ('C:\\Users\\Bailey\\PycharmProjects\\Vce_Viewer\\.venv\\Lib\\site-packages\\PyQt6\\Qt6\\translations', 'PyQt6/Qt6/translations'), ('C:\\Users\\Bailey\\PycharmProjects\\VCE_Viewer\\uploaded_reports', 'uploaded_reports')],
    hiddenimports=['PyQt6.QtPdf', 'PyQt6.QtPdfWidgets', 'bs4', 'requests'],
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
    name='VCEViewer',
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
    icon=['C:\\Users\\Bailey\\PycharmProjects\\VCE_Viewer\\app_icon.ico'],
)
