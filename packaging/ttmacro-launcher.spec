# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec：GUI ランチャー（ttmacro-launcher.exe）。

ビルド方法（プロジェクトルートで PowerShell）:
    .venv\\Scripts\\pyinstaller --clean --distpath bin packaging/ttmacro-launcher.spec

成果物は bin/ttmacro-launcher.exe。--onefile で単一 exe として固める。
GUI なので console=False（実行時に黒いコンソール窓が出ない）。
"""

from pathlib import Path

# .spec ファイルからプロジェクトルートを解決（PyInstaller が SPECPATH を提供）
PROJECT_ROOT = Path(SPECPATH).resolve().parent  # noqa: F821 (SPECPATH は PyInstaller 注入)
SRC_DIR = PROJECT_ROOT / "src"

# アイコン未用意のため None。.ico ファイル用意後に有効化する。
# 例: ICON_PATH = str(PROJECT_ROOT / "packaging" / "launcher.ico")
ICON_PATH = None

a = Analysis(
    [str(PROJECT_ROOT / "packaging" / "entry_launcher.py")],
    pathex=[str(SRC_DIR)],
    binaries=[],
    datas=[],
    # ttkbootstrap は実行時にテーマファイルを動的読込するため hidden import 指定で確実に拾う
    hiddenimports=["ttkbootstrap"],
    hookspath=[],
    runtime_hooks=[],
    excludes=[
        # 不要なテスト系/巨大ライブラリは除外してサイズを抑える
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
    name="ttmacro-launcher",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI モード：コンソール窓を出さない
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON_PATH,
)
