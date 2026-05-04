"""Excel 台帳の読み込み・行検証・行データ抽出。

このモジュールは pandas に依存するため、Python 3.14 環境では
import で失敗する可能性がある。呼び出し元（cli.py）は遅延 import で
このリスクをユーザーフレンドリーなメッセージに変換する。
"""

from __future__ import annotations

import ipaddress
import math
import re

import pandas as pd

from ttmacro.config import EXCEL_PATH, KEYS_DIR
from ttmacro.ttl_renderer import sanitize_name


def safe_str(val: object) -> str:
    """Excel の NaN を空文字に変換する。

    Args:
        val: Excel セルから読んだ生の値。

    Returns:
        NaN なら空文字、それ以外は str() を strip した文字列。
    """
    if isinstance(val, float) and math.isnan(val):
        return ""
    return str(val).strip()


def safe_get(row: pd.Series, key: str, default: str = "") -> str:
    """行データから安全に値を取得し、NaN なら default を返す。

    Args:
        row: pandas Series（1行分の Excel データ）。
        key: 取得するカラム名。
        default: NaN 時のフォールバック値。

    Returns:
        値の文字列表現（strip 済み）。
    """
    value = row.get(key, default)
    return str(value if pd.notna(value) else default).strip()


def load_excel_data() -> pd.DataFrame:
    """Excel 台帳ファイルを読み込む。

    Returns:
        Excel から読み込んだ DataFrame。

    Raises:
        FileNotFoundError: ファイルが存在しない場合。
        PermissionError: ファイルが他のアプリで開かれている場合。
        ValueError: ファイルが空の場合。
        RuntimeError: その他の読み込みエラー。
    """
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {EXCEL_PATH}")

    try:
        with open(EXCEL_PATH, "rb") as f:
            df = pd.read_excel(f, engine="openpyxl")
            if df.empty:
                raise ValueError("Excelファイルが空です")
            return df
    except PermissionError as e:
        raise PermissionError(
            f"Excelファイルが他で開かれています: {EXCEL_PATH}"
        ) from e
    except (FileNotFoundError, ValueError):
        # 上で投げた例外はそのまま伝播
        raise
    except Exception as e:
        raise RuntimeError(f"Excelファイル読み込みエラー: {e}") from e


def validate_row_data(row: pd.Series, row_num: int) -> tuple[bool, list[str]]:
    """行データの妥当性を検証する。

    必須項目（name/host/user）、ホスト名/IP の形式、ポート番号の範囲、
    keyfile の存在をチェックする。

    Args:
        row: 検証対象の行（pandas Series）。
        row_num: ログ用の行番号（現状は未使用、将来用に保持）。

    Returns:
        ``(is_valid, errors)`` のタプル。
    """
    errors: list[str] = []

    # 必須フィールドチェック
    required_fields = ["name", "host", "user"]
    for field in required_fields:
        if pd.isna(row.get(field)) or str(row.get(field, "")).strip() == "":
            errors.append(f"必須項目 '{field}' が空です")

    # IPアドレス/ホスト名チェック
    host = str(row.get("host", "")).strip()
    if host:
        try:
            ipaddress.ip_address(host)
        except ValueError:
            # IP として無効ならホスト名扱い（簡易チェック）
            if not re.match(r"^[a-zA-Z0-9.-]+$", host):
                errors.append(f"ホスト名 '{host}' の形式が不正です")

    # ポート番号チェック
    port = row.get("port")
    if pd.notna(port):
        try:
            port_num = int(port)
            if not (1 <= port_num <= 65535):
                errors.append(f"ポート番号 {port_num} は範囲外です (1-65535)")
        except (ValueError, TypeError):
            errors.append(f"ポート番号 '{port}' が数値ではありません")

    # キーファイル存在チェック
    keyfile = safe_get(row, "keyfile")
    if keyfile:
        keyfile_path = KEYS_DIR / keyfile
        if not keyfile_path.exists():
            errors.append(
                f"キーファイル '{keyfile}' が見つかりません: {keyfile_path}"
            )

    return len(errors) == 0, errors


def extract_row_data(row: pd.Series) -> dict[str, str]:
    """行データから TTL 生成に必要な情報を抽出する。

    Args:
        row: 抽出元の行（pandas Series）。

    Returns:
        name/host/port/user/password/keyfile_name/post_cmd/memo/group1-3 を
        含む辞書。
    """
    # メモ内の改行・タブは半角空白に置換（TTL コメントが壊れないように）
    memo = (
        safe_get(row, "memo")
        .replace("\r", " ")
        .replace("\n", " ")
        .replace("\t", " ")
    )

    return {
        "name": sanitize_name(str(row["name"]).strip()),
        "host": str(row["host"]).strip(),
        "port": str(int(row["port"])) if pd.notna(row["port"]) else "22",
        "user": str(row["user"]).strip(),
        "password": safe_get(row, "password"),
        "keyfile_name": safe_get(row, "keyfile"),
        "post_cmd": safe_get(row, "post_cmd"),
        "memo": memo,
        "group1": safe_get(row, "group1"),
        "group2": safe_get(row, "group2"),
        "group3": safe_get(row, "group3"),
    }
