# -*- mode: python ; coding: utf-8 -*-
"""Cross-platform PyInstaller spec for WordCombine."""

import os
import sys

ROOT = os.path.dirname(os.path.abspath(SPECPATH))

hidden_imports = [
    "lxml.etree",
    "lxml._elementpath",
]
if sys.platform == "win32":
    hidden_imports.extend(
        [
            "win32com",
            "win32com.client",
            "pythoncom",
            "pywintypes",
            "win32timezone",
        ]
    )

excluded_modules = [
    "matplotlib",
    "numpy",
    "pandas",
    "scipy",
    "PIL",
    "IPython",
    "jedi",
    "streamlit",
    "fastapi",
    "uvicorn",
    "openai",
    "google",
    "grpc",
    "pyarrow",
    "reportlab",
    "pywebview",
    "pytest",
    "docx",
]

datas = [
    (os.path.join(ROOT, "assets", "icon.png"), "assets"),
    (os.path.join(ROOT, "assets", "icon.ico"), "assets"),
]

a = Analysis(
    [os.path.join(ROOT, "word_combine.py")],
    pathex=[ROOT],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excluded_modules,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

if sys.platform == "win32":
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name="WordCombine",
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
        icon=os.path.join(ROOT, "assets", "icon.ico"),
    )
elif sys.platform == "darwin":
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name="WordCombine",
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
        icon=os.path.join(ROOT, "assets", "icon.icns"),
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        upx_exclude=[],
        name="WordCombine",
    )
    app = BUNDLE(
        coll,
        name="WordCombine.app",
        icon=os.path.join(ROOT, "assets", "icon.icns"),
        bundle_identifier="com.artxom.wordcombine",
        version="1.0.0",
        info_plist={
            "CFBundleShortVersionString": "1.0.0",
            "CFBundleVersion": "1.0.0",
            "NSHighResolutionCapable": True,
        },
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name="WordCombine",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=False,
    )
