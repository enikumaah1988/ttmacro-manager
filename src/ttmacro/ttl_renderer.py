"""TTL テンプレート読み込みとプレースホルダ展開。

`{name}` `{hostname}` `{port}` 等のプレースホルダを行データで置換する。
str.replace のループによる単純置換のため、置換値（例: memo）に
``{password}`` のような文字列が含まれると誤置換される。
（既存仕様維持。将来的に Jinja2 等への置換を検討）
"""

from __future__ import annotations

import re
from pathlib import Path

from ttmacro.config import TEMPLATE_PATH, TEMPLATES_DIR
from ttmacro.path_resolver import calculate_relative_path


def sanitize_name(name: str) -> str:
    """Windows のファイル名禁止文字を ``_`` に置換する。

    Args:
        name: 任意の文字列。

    Returns:
        ``\\/:*?"<>|`` を ``_`` に置換した文字列。
    """
    return re.sub(r'[\\/:*?"<>|]', "_", name)


def resolve_template_path(template_name: str) -> Path:
    """Excel の ``template`` 列値からテンプレートファイルのパスを返す。

    空または未指定（``""``）なら ``TEMPLATE_PATH``（デフォルト）を返す。
    値が指定されていれば ``TEMPLATES_DIR / <name>.ttl`` を返す。
    拡張子 ``.ttl`` の有無は許容。サブパス（例: ``"subdir/foo"``）も可。

    パストラバーサル防止のため、結果が ``TEMPLATES_DIR`` の外を指している
    場合は ``ValueError`` を投げる。実在チェックは行わない。

    Args:
        template_name: Excel の ``template`` 列の値。

    Returns:
        テンプレートファイルの絶対パス。

    Raises:
        ValueError: パスが ``TEMPLATES_DIR`` の外を指している場合。
    """
    if not template_name:
        return TEMPLATE_PATH

    name = template_name.strip()
    if name.endswith(".ttl"):
        name = name[:-4]

    resolved = (TEMPLATES_DIR / f"{name}.ttl").resolve()
    templates_root = TEMPLATES_DIR.resolve()
    if not resolved.is_relative_to(templates_root):
        raise ValueError(
            f"テンプレート '{template_name}' が TEMPLATES_DIR の外を指しています: {resolved}"
        )
    return resolved


def load_template(path: Path | None = None) -> str:
    """指定されたテンプレートファイルを読み込む。

    Args:
        path: 読み込むテンプレートのパス。``None`` ならデフォルトの
            ``TEMPLATE_PATH`` を使う（後方互換）。

    Returns:
        テンプレート本文の文字列。

    Raises:
        FileNotFoundError: テンプレートファイルが存在しない場合。
        ValueError: 空ファイル、または UTF-8 として読めない場合。
        RuntimeError: その他の読み込みエラー。
    """
    target = path if path is not None else TEMPLATE_PATH

    if not target.exists():
        raise FileNotFoundError(f"テンプレートファイルが見つかりません: {target}")

    try:
        content = target.read_text(encoding="utf-8")
        if not content.strip():
            raise ValueError(f"テンプレートファイルが空です: {target}")
        return content
    except UnicodeDecodeError as e:
        raise ValueError(
            f"テンプレートファイルの文字エンコーディングが不正です: {target}"
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
