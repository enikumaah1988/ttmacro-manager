# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec：TTL 生成 CLI（ttmacro-generate.exe）。

ビルド方法（プロジェクトルートで PowerShell）:
    .venv\\Scripts\\pyinstaller --clean --distpath bin packaging/ttmacro-generate.spec

成果物は bin/ttmacro-generate.exe。--onefile で単一 exe として固める。
CLI なので console=True（コンソールに進捗・ログを出す）。
"""

from pathlib import Path

PROJECT_ROOT = Path(SPECPATH).resolve().parent  # noqa: F821
SRC_DIR = PROJECT_ROOT / "src"

# ランチャーと同じアイコンを使う
ICON_PATH = str(PROJECT_ROOT / "packaging" / "launcher.ico")

a = Analysis(
    [str(PROJECT_ROOT / "packaging" / "entry_generate.py")],
    pathex=[str(SRC_DIR)],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        # ランチャー専用の GUI 系は CLI exe に同梱しない（サイズ削減）
        "ttkbootstrap",
        "tkinter",
        # その他不要な巨大ライブラリ
        "pandas",
        "numpy",
        "pytest",
        "mypy",
        "ruff",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="ttmacro-generate",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # CLI モード：コンソールに出力
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON_PATH,
)
