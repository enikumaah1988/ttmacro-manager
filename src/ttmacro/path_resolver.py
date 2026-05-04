"""TTL ファイルの出力先・相対パス計算。

副作用（mkdir、書き込み権限テスト）を伴う関数と、
純粋なパス計算関数が混在している（既存仕様に合わせて保持）。
"""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

from ttmacro.config import BASE_DIR, KEYS_DIR, LOGS_DIR, OUTPUT_DIR


def get_target_directory(data: dict[str, str]) -> Path:
    """グループ階層に基づいて出力ディレクトリを決定する。

    副作用として、ディレクトリ作成と書き込み権限テストを行う。
    ``group1`` が空欄の場合は ``OUTPUT_DIR`` 直下を返す（``group2`` /
    ``group3`` だけの指定は無効）。

    Args:
        data: ``extract_row_data`` が返す行データ辞書。

    Returns:
        作成済み（または既存）の出力ディレクトリ。

    Raises:
        PermissionError: ディレクトリへの書き込み権限がない場合。
        RuntimeError: その他のディレクトリ作成エラー。
    """
    if not data["group1"]:
        return OUTPUT_DIR

    target_dir = OUTPUT_DIR / data["group1"]
    if data["group2"]:
        target_dir = target_dir / data["group2"]
        if data["group3"]:
            target_dir = target_dir / data["group3"]

    try:
        target_dir.mkdir(parents=True, exist_ok=True)
        # 書き込み権限チェック（一時ファイルを作って消す）
        test_file = target_dir / ".write_test"
        test_file.touch()
        test_file.unlink()
    except PermissionError as e:
        raise PermissionError(
            f"ディレクトリへの書き込み権限がありません: {target_dir}"
        ) from e
    except Exception as e:
        raise RuntimeError(f"ディレクトリ作成エラー: {target_dir} - {e}") from e

    return target_dir


def calculate_relative_path(target_dir: Path) -> str:
    """TTL 配置場所からプロジェクトルートへの相対パスを計算する。

    TTL 内で ``getdir`` と組み合わせてプロジェクトルートを特定するために使う。

    Args:
        target_dir: TTL 出力先ディレクトリの絶対パス。

    Returns:
        ``'../'`` を必要数連結した文字列。ルート直下なら空文字。
    """
    rel_path = target_dir.relative_to(BASE_DIR)

    if rel_path == Path("."):
        return ""

    depth = len(rel_path.parts)
    return "../" * depth


def get_log_dir(target_dir: Path) -> Path:
    """TTL と同じ階層になるよう logs 以下のディレクトリを返す。

    Example:
        ``macros/home/prod`` → ``logs/home/prod``
    """
    rel = target_dir.relative_to(OUTPUT_DIR)
    if rel == Path("."):
        return LOGS_DIR
    return LOGS_DIR / rel


def calculate_paths(data: dict[str, str], target_dir: Path) -> dict[str, str]:
    """各種パスを計算する（現状は呼び出し元で結果が利用されていない）。

    既存仕様維持のため残置。リファクタ完了後の整理対象。

    Args:
        data: 行データ辞書。
        target_dir: 出力先ディレクトリ。

    Returns:
        ttl_name / log_file / log_path / keyfile を含む辞書。
    """
    ttl_name = f"{data['name']}_{data['host']}_{data['user']}"

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"{ttl_name}_{timestamp}.log"
    log_file = LOGS_DIR / log_filename

    keyfile_path = ""
    if data["keyfile_name"]:
        keyfile_path = str(KEYS_DIR / data["keyfile_name"])

    return {
        "ttl_name": ttl_name,
        "log_file": str(log_file),
        "log_path": str(LOGS_DIR),
        "keyfile": keyfile_path,
    }
