"""パス定数の集約モジュール。

プロジェクトルート起点の各種パスを集約する。利用側は `from ttmacro.config import EXCEL_PATH` のように参照する。
"""

from __future__ import annotations

from pathlib import Path

# src/ttmacro/config.py からプロジェクトルートまで 3 階層上る
BASE_DIR = Path(__file__).resolve().parent.parent.parent
EXCEL_PATH = BASE_DIR / "data" / "servers.xlsx"
TEMPLATE_PATH = BASE_DIR / "macros" / "template.ttl"
OUTPUT_DIR = BASE_DIR / "macros"
LOGS_DIR = BASE_DIR / "logs"
KEYS_DIR = BASE_DIR / "keys"
