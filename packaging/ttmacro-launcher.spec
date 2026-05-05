# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec：GUI ランチャー（ttmacro-launcher）。

ビルド方法（プロジェクトルートで PowerShell）:
    .venv\\Scripts\\pyinstaller --clean --distpath bin packaging/ttmacro-launcher.spec

成果物は **bin/ttmacro-launcher/ フォルダ** で、その中に
``ttmacro-launcher.exe`` と依存ファイル一式が入る（``--onedir`` モード）。
``--onefile`` では起動時に %TEMP% へ全リソースを展開するため Windows Defender の
スキャンが入って 10〜15 秒かかっていたが、``--onedir`` なら展開が不要なため
1 秒前後で起動する。UPX 圧縮も解凍コストを避けるため無効化している。

GUI なので console=False（実行時に黒いコンソール窓が出ない）。
"""

from pathlib import Path

# .spec ファイルからプロジェクトルートを解決（PyInstaller が SPECPATH を提供）
PROJECT_ROOT = Path(SPECPATH).resolve().parent  # noqa: F821 (SPECPATH は PyInstaller 注入)
SRC_DIR = PROJECT_ROOT / "src"

ICON_PATH = str(PROJECT_ROOT / "packaging" / "launcher.ico")

a = Analysis(
    [str(PROJECT_ROOT / "packaging" / "entry_launcher.py")],
    pathex=[str(SRC_DIR)],
    binaries=[],
    # ウィンドウ／タスクバーアイコン用に .ico を成果物フォルダへ同梱
    datas=[
        (str(PROJECT_ROOT / "packaging" / "launcher.ico"), "."),
    ],
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
    [],
    exclude_binaries=True,  # --onedir：バイナリを exe に内包せず COLLECT で配置
    name="ttmacro-launcher",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # 起動時の UPX 解凍を避ける（速度優先）
    console=False,  # GUI モード：コンソール窓を出さない
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON_PATH,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ttmacro-launcher",
)
