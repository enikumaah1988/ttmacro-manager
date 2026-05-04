"""TTL テンプレート読み込みとプレースホルダ展開。

`{name}` `{hostname}` `{port}` 等のプレースホルダを行データで置換する。
str.replace のループによる単純置換のため、置換値（例: memo）に
``{password}`` のような文字列が含まれると誤置換される。
（既存仕様維持。将来的に Jinja2 等への置換を検討）
"""

from __future__ import annotations

import re
from pathlib import Path

from ttmacro.config import TEMPLATE_PATH
from ttmacro.path_resolver import calculate_relative_path


def sanitize_name(name: str) -> str:
    """Windows のファイル名禁止文字を ``_`` に置換する。

    Args:
        name: 任意の文字列。

    Returns:
        ``\\/:*?"<>|`` を ``_`` に置換した文字列。
    """
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def load_template() -> str:
    """TTL テンプレートを読み込む。

    Returns:
        テンプレート本文の文字列。

    Raises:
        FileNotFoundError: テンプレートファイルが存在しない場合。
        ValueError: 空ファイル、または UTF-8 として読めない場合。
        RuntimeError: その他の読み込みエラー。
    """
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"テンプレートファイルが見つかりません: {TEMPLATE_PATH}"
        )

    try:
        content = TEMPLATE_PATH.read_text(encoding="utf-8")
        if not content.strip():
            raise ValueError("テンプレートファイルが空です")
        return content
    except UnicodeDecodeError as e:
        raise ValueError(
            f"テンプレートファイルの文字エンコーディングが不正です: {TEMPLATE_PATH}"
        ) from e
    except (FileNotFoundError, ValueError):
        # 上で投げた例外はそのまま伝播
        raise
    except Exception as e:
        raise RuntimeError(f"テンプレートファイル読み込みエラー: {e}") from e


def generate_ttl_content(
    data: dict[str, str], template: str, timestamp: str, target_dir: Path
) -> str:
    """TTL マクロ本文を生成する。

    Args:
        data: ``extract_row_data`` が返す行データ辞書。
        template: ``load_template()`` で読み込んだテンプレート本文。
        timestamp: テンプレートの ``{created_at}`` に埋め込む文字列。
        target_dir: TTL 出力先ディレクトリ（相対パス計算に使用）。

    Returns:
        プレースホルダを置換済みの TTL 本文。
    """
    rel_path = calculate_relative_path(target_dir)

    # ポストコマンドを TTL 構文の wait/sendln ペアに展開
    post_cmd_lines = [
        line.strip() for line in data["post_cmd"].splitlines() if line.strip()
    ]
    post_commands = (
        "\n".join([f"wait '$' '#'\nsendln '{cmd}'\n" for cmd in post_cmd_lines])
        if post_cmd_lines
        else ""
    )

    replacements = {
        "{hostname}": data["host"],
        "{port}": data["port"],
        "{username}": data["user"],
        "{password}": data["password"],
        "{keyfile}": data["keyfile_name"],  # キーファイル名のみ（パスは TTL 側で合成）
        "{name}": data["name"],
        "{rel_path}": rel_path,
        "{created_at}": timestamp,
        "{memo}": data["memo"],
        "{post_commands}": post_commands,
    }

    content = template
    for key, value in replacements.items():
        content = content.replace(key, value)

    return content
