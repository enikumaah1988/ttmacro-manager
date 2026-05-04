"""パス定数の集約モジュール。

プロジェクトルート起点の各種パスを集約する。利用側は
``from ttmacro.config import EXCEL_PATH`` のように参照する。

各パスは以下の環境変数で上書き可能：

- ``TTMACRO_BASE_DIR``: プロジェクトルート（他のパスのデフォルト起点）
- ``TTMACRO_EXCEL_PATH``: Excel 台帳ファイル
- ``TTMACRO_TEMPLATE_PATH``: デフォルトテンプレート（Excel template 列が空の時）
- ``TTMACRO_TEMPLATES_DIR``: 追加テンプレート格納ディレクトリ
- ``TTMACRO_OUTPUT_DIR``: TTL 生成先ディレクトリ
- ``TTMACRO_LOGS_DIR``: ログ出力先ディレクトリ
- ``TTMACRO_KEYS_DIR``: 鍵ファイル格納ディレクトリ

優先順位は「個別の環境変数 > ``TTMACRO_BASE_DIR`` 起点 > リポジトリ既定値」。
環境変数はモジュール ``import`` 時に 1 回だけ評価される。実行中に環境変数を
変えても本モジュールの定数は変化しない。
"""

from __future__ import annotations

import os
import sys
from pathlib import Path


def _env_path(env_name: str, default: Path) -> Path:
    """環境変数 ``env_name`` を Path として読み取り、未設定または空文字なら
    ``default`` を返す。

    Args:
        env_name: 参照する環境変数名。
        default: 環境変数が無い場合のフォールバックパス。

    Returns:
        環境変数値を ``Path.resolve()`` した絶対パス、または ``default``。
    """
    value = os.environ.get(env_name)
    if value:
        return Path(value).resolve()
    return default


def _resolve_default_base_dir() -> Path:
    """デフォルトの BASE_DIR を実行モード別に解決する。

    開発実行（python -m / pip install -e .）では本ファイルから
    プロジェクトルートまで 3 階層上った場所を返す。

    PyInstaller で凍結された実行ファイルでは ``sys.executable`` が
    ``<deploy_root>/bin/<exe>`` を指すため、その親の親をデプロイルート
    として返す（同梱の data/macros/keys/logs を相対参照するため）。
    """
    if getattr(sys, "frozen", False):
        # PyInstaller --onefile：sys.executable は bin/<exe>。
        # その親（bin/）の親 = デプロイルート。
        return Path(sys.executable).resolve().parent.parent
    # 開発実行：src/ttmacro/config.py から 3 階層上る
    return Path(__file__).resolve().parent.parent.parent


_DEFAULT_BASE_DIR = _resolve_default_base_dir()

# 個別の環境変数 > BASE_DIR 起点デフォルト の優先順位
BASE_DIR = _env_path("TTMACRO_BASE_DIR", _DEFAULT_BASE_DIR)
EXCEL_PATH = _env_path("TTMACRO_EXCEL_PATH", BASE_DIR / "data" / "servers.xlsx")
TEMPLATES_DIR = _env_path("TTMACRO_TEMPLATES_DIR", BASE_DIR / "macros" / "templates")
# デフォルトテンプレートは TEMPLATES_DIR/default.ttl を採用。
# Excel の template 列が空の行はこのファイルが使われる。
TEMPLATE_PATH = _env_path("TTMACRO_TEMPLATE_PATH", TEMPLATES_DIR / "default.ttl")
OUTPUT_DIR = _env_path("TTMACRO_OUTPUT_DIR", BASE_DIR / "macros")
LOGS_DIR = _env_path("TTMACRO_LOGS_DIR", BASE_DIR / "logs")
KEYS_DIR = _env_path("TTMACRO_KEYS_DIR", BASE_DIR / "keys")
