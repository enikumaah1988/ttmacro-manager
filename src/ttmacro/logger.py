"""ttmacro パッケージのロギング設定。

`setup_logging()` を一度呼ぶことで、'ttmacro' 親ロガーにファイル/コンソール
の両ハンドラを設定する。各モジュールは `logging.getLogger(__name__)` で
ロガーを取得すれば、親に伝播してハンドラが利用される。
"""

from __future__ import annotations

import logging
import sys

from ttmacro.config import LOGS_DIR


def setup_logging() -> logging.Logger:
    """'ttmacro' ロガーにハンドラを設定して返す。

    ファイル出力先は ``logs/generate.log``、フォーマットは
    ``[YYYY-MM-DD HH:MM:SS] message``。コンソールは stderr。

    Returns:
        設定済みの 'ttmacro' ロガー。
    """
    log_file = LOGS_DIR / "generate.log"
    LOGS_DIR.mkdir(exist_ok=True)

    formatter = logging.Formatter(
        "[%(asctime)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S"
    )
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setFormatter(formatter)

    logger = logging.getLogger("ttmacro")
    # 再 setup 時にハンドラが重複しないようクリアしてから追加
    logger.handlers.clear()
    logger.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    # ルートロガーへの伝播を止める（親側の設定で二重出力されるのを防ぐ）
    logger.propagate = False
    return logger
