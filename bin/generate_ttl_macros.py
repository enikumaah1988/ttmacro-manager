"""TTL マクロ生成スクリプトのエントリポイント（薄いラッパ）。

実装は src/ttmacro/cli.py に移動済み。このファイルは後方互換のため、
従来通り ``python bin/generate_ttl_macros.py`` で実行可能にする。

pip install していない環境でも動くよう、src/ を sys.path に追加してから
ttmacro.cli.main を呼び出す。
"""

from __future__ import annotations

import sys
from pathlib import Path

# pandas インポート前に Python のパスを表示（3.14 環境で pandas が落ちる場合の調査支援）
print("使用中の Python:", sys.executable, file=sys.stderr, flush=True)
print("TTLマクロ生成を開始しています...", file=sys.stderr, flush=True)

# src/ をパスに追加（pip install -e . していなくても直接実行できるように）
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_PROJECT_ROOT / "src"))

from ttmacro.cli import main  # noqa: E402  sys.path 設定後に import する必要がある


if __name__ == "__main__":
    main()
