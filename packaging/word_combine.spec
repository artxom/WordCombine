# -*- mode: python ; coding: utf-8 -*-
"""Cross-platform PyInstaller spec for WordCombine."""

import os
import re
import sys

ROOT = os.path.dirname(os.path.abspath(SPECPATH))


def _read_app_version() -> str:
    init_path = os.path.join(ROOT, "word_combine", "__init__.py")
    try:
        with open(init_path, "r", encoding="utf-8") as fh:
            match = re.search(r'^__version__\s*=\s*"([^"]+)"', fh.read(), re.MULTILINE)
            if match:
                return match.group(1)
    except OSError:
        pass
    return "1.0.1"


APP_VERSION = _read_app_version()

hidden_imports = [
    "lxml.etree",
    "lxml._elementpath",
    "word_combine.ooxml_merger",
    "word_combine.toc_refresher",
    "word_combine.verifier",
    "word_combine.cli",
    "word_combine.gui",
    "word_combine.sorting",
    "word_combine.splash",
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


def _keep_data_entry(entry) -> bool:
    dest = entry[0].replace("\\", "/").lower()
    if dest.startswith(
        (
            "_tcl_data/tzdata/",
            "_tcl_data/http1.0/",
            "_tcl_data/opt0.4/",
            "_tk_data/images/",
            "_tk_data/demos/",
        )
    ):
        return False
    for msg_prefix in ("_tcl_data/msgs/", "_tk_data/msgs/"):
        if dest.startswith(msg_prefix):
            base = dest[len(msg_prefix) :]
            if not base.startswith(("en", "zh")):
                return False
    return True


a.datas[:] = [entry for entry in a.datas if _keep_data_entry(entry)]

pyz = PYZ(a.pure)

if sys.platform == "win32":
    win_mode = os.environ.get("WORDCOMBINE_WIN_MODE", "onefile").strip().lower()
    if win_mode == "onedir":
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
            icon=os.path.join(ROOT, "assets", "icon.ico"),
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
    else:
        splash = Splash(
            os.path.join(ROOT, "assets", "splash.png"),
            binaries=a.binaries,
            datas=a.datas,
            text_pos=(28, 180),
            text_size=9,
            text_color="#64748B",
            text_default="Starting WordCombine...",
            minify_script=True,
            always_on_top=True,
            center="active",
        )
        exe = EXE(
            pyz,
            a.scripts,
            splash,
            splash.binaries,
            a.binaries,
            a.datas,
            [],
            name="WordCombine",
            debug=False,
            bootloader_ignore_signals=False,
            strip=False,
            upx=False,
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
        version=APP_VERSION,
        info_plist={
            "CFBundleShortVersionString": APP_VERSION,
            "CFBundleVersion": APP_VERSION,
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
