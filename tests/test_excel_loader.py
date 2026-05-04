"""ttmacro.excel_loader のテスト。

openpyxl ベース実装に移行済み。1 行は ``dict[str, Any]`` で表現する。
空セルは ``None`` で表すのが原則（Excel 由来の値の再現）。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import pytest

from ttmacro.excel_loader import (
    extract_row_data,
    is_blank,
    is_blank_row,
    safe_get,
    safe_str,
    validate_row_data,
)


class TestIsBlank:
    """is_blank のテスト。"""

    def test_none_is_blank(self) -> None:
        assert is_blank(None)

    def test_empty_string_is_blank(self) -> None:
        assert is_blank("")

    def test_whitespace_only_is_blank(self) -> None:
        assert is_blank("   ")

    def test_nan_is_blank(self) -> None:
        assert is_blank(float("nan"))

    def test_normal_string_is_not_blank(self) -> None:
        assert not is_blank("x")

    def test_zero_is_not_blank(self) -> None:
        assert not is_blank(0)


class TestIsBlankRow:
    """is_blank_row のテスト。"""

    def test_all_none_is_blank(self) -> None:
        assert is_blank_row({"a": None, "b": None})

    def test_mixed_blanks(self) -> None:
        assert is_blank_row({"a": None, "b": "", "c": "  "})

    def test_one_value_is_not_blank(self) -> None:
        assert not is_blank_row({"a": None, "b": "x"})


class TestSafeStr:
    """safe_str のテスト。"""

    def test_normal_string(self) -> None:
        assert safe_str("hello") == "hello"

    def test_strips_whitespace(self) -> None:
        assert safe_str("  hello  ") == "hello"

    def test_nan_returns_empty(self) -> None:
        assert safe_str(float("nan")) == ""

    def test_none_returns_empty(self) -> None:
        assert safe_str(None) == ""

    def test_integer_value(self) -> None:
        assert safe_str(42) == "42"

    def test_empty_string(self) -> None:
        assert safe_str("") == ""


class TestSafeGet:
    """safe_get のテスト。"""

    def test_normal_value(self) -> None:
        row: dict[str, Any] = {"key": "  value  "}
        assert safe_get(row, "key") == "value"

    def test_none_returns_default(self) -> None:
        row: dict[str, Any] = {"key": None}
        assert safe_get(row, "key", "fallback") == "fallback"

    def test_missing_key_returns_default(self) -> None:
        row: dict[str, Any] = {"other": "x"}
        assert safe_get(row, "key", "fallback") == "fallback"

    def test_missing_key_no_default_returns_empty(self) -> None:
        row: dict[str, Any] = {"other": "x"}
        assert safe_get(row, "key") == ""


class TestValidateRowData:
    """validate_row_data のテスト。"""

    @pytest.fixture
    def valid_row(self) -> dict[str, Any]:
        """検証を通過する基本的な行。"""
        return {
            "name": "srv01",
            "host": "192.168.0.10",
            "user": "admin",
            "port": 22,
            "keyfile": "",
        }

    def test_valid_row_passes(self, valid_row: dict[str, Any]) -> None:
        is_valid, errors = validate_row_data(valid_row, 1)
        assert is_valid
        assert errors == []

    def test_missing_name_fails(self, valid_row: dict[str, Any]) -> None:
        valid_row["name"] = ""
        is_valid, errors = validate_row_data(valid_row, 1)
        assert not is_valid
        assert any("name" in e for e in errors)

    def test_missing_host_fails(self, valid_row: dict[str, Any]) -> None:
        valid_row["host"] = ""
        is_valid, errors = validate_row_data(valid_row, 1)
        assert not is_valid
        assert any("host" in e for e in errors)

    def test_missing_user_fails(self, valid_row: dict[str, Any]) -> None:
        valid_row["user"] = ""
        is_valid, errors = validate_row_data(valid_row, 1)
        assert not is_valid
        assert any("user" in e for e in errors)

    def test_hostname_format_accepted(self, valid_row: dict[str, Any]) -> None:
        """ホスト名（IP でない形式）が許容される。"""
        valid_row["host"] = "server01.example.com"
        is_valid, errors = validate_row_data(valid_row, 1)
        assert is_valid
        assert errors == []

    def test_invalid_hostname_format_fails(self, valid_row: dict[str, Any]) -> None:
        """記号を含む不正なホスト名は弾かれる。"""
        valid_row["host"] = "invalid host!"
        is_valid, errors = validate_row_data(valid_row, 1)
        assert not is_valid
        assert any("ホスト名" in e for e in errors)

    def test_port_out_of_range_fails(self, valid_row: dict[str, Any]) -> None:
        valid_row["port"] = 70000
        is_valid, errors = validate_row_data(valid_row, 1)
        assert not is_valid
        assert any("ポート番号" in e for e in errors)

    def test_port_zero_fails(self, valid_row: dict[str, Any]) -> None:
        valid_row["port"] = 0
        is_valid, errors = validate_row_data(valid_row, 1)
        assert not is_valid

    def test_port_non_numeric_fails(self, valid_row: dict[str, Any]) -> None:
        valid_row["port"] = "abc"
        is_valid, errors = validate_row_data(valid_row, 1)
        assert not is_valid

    def test_keyfile_missing_fails(
        self,
        valid_row: dict[str, Any],
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """指定された keyfile が KEYS_DIR に無い場合エラー。"""
        monkeypatch.setattr("ttmacro.excel_loader.KEYS_DIR", tmp_path)
        valid_row["keyfile"] = "missing.key"
        is_valid, errors = validate_row_data(valid_row, 1)
        assert not is_valid
        assert any("キーファイル" in e for e in errors)

    def test_keyfile_existing_passes(
        self,
        valid_row: dict[str, Any],
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
    ) -> None:
        """keyfile が KEYS_DIR に存在する場合は通過。"""
        monkeypatch.setattr("ttmacro.excel_loader.KEYS_DIR", tmp_path)
        (tmp_path / "id.key").touch()
        valid_row["keyfile"] = "id.key"
        is_valid, errors = validate_row_data(valid_row, 1)
        assert is_valid


class TestExtractRowData:
    """extract_row_data のテスト。"""

    def test_basic_conversion(self) -> None:
        row: dict[str, Any] = {
            "name": "  srv01  ",
            "host": "10.0.0.1",
            "port": 2222,
            "user": "admin",
            "password": "pass",
            "keyfile": "",
            "post_cmd": "ls",
            "memo": "test server",
            "group1": "envA",
            "group2": "",
            "group3": "",
        }
        data = extract_row_data(row)
        assert data["name"] == "srv01"
        assert data["host"] == "10.0.0.1"
        assert data["port"] == "2222"
        assert data["user"] == "admin"
        assert data["password"] == "pass"
        assert data["post_cmd"] == "ls"
        assert data["memo"] == "test server"
        assert data["group1"] == "envA"

    def test_missing_port_defaults_to_22(self) -> None:
        row: dict[str, Any] = {
            "name": "srv",
            "host": "10.0.0.1",
            "port": None,
            "user": "admin",
            "password": "",
            "keyfile": "",
            "post_cmd": "",
            "memo": "",
            "group1": "",
            "group2": "",
            "group3": "",
        }
        data = extract_row_data(row)
        assert data["port"] == "22"

    def test_memo_with_newlines_normalized_to_spaces(self) -> None:
        """メモの改行・タブが半角空白に置換される（TTL コメントを壊さないため）。"""
        row: dict[str, Any] = {
            "name": "srv",
            "host": "10.0.0.1",
            "port": 22,
            "user": "admin",
            "password": "",
            "keyfile": "",
            "post_cmd": "",
            "memo": "line1\nline2\rline3\tline4",
            "group1": "",
            "group2": "",
            "group3": "",
        }
        data = extract_row_data(row)
        assert data["memo"] == "line1 line2 line3 line4"

    def test_name_sanitized(self) -> None:
        """name に Windows 禁止文字が含まれていればサニタイズされる。"""
        row: dict[str, Any] = {
            "name": "srv:01/test",
            "host": "10.0.0.1",
            "port": 22,
            "user": "admin",
            "password": "",
            "keyfile": "",
            "post_cmd": "",
            "memo": "",
            "group1": "",
            "group2": "",
            "group3": "",
        }
        data = extract_row_data(row)
        assert data["name"] == "srv_01_test"

    def test_template_field_extracted(self) -> None:
        """template 列の値が抽出される。"""
        row: dict[str, Any] = {
            "name": "srv",
            "host": "10.0.0.1",
            "port": 22,
            "user": "admin",
            "password": "",
            "keyfile": "",
            "post_cmd": "",
            "memo": "",
            "group1": "",
            "group2": "",
            "group3": "",
            "template": "  nodejs  ",
        }
        data = extract_row_data(row)
        assert data["template"] == "nodejs"

    def test_template_missing_defaults_to_empty(self) -> None:
        """template 列が無い場合、空文字になる（既存 Excel への後方互換）。"""
        row: dict[str, Any] = {
            "name": "srv",
            "host": "10.0.0.1",
            "port": 22,
            "user": "admin",
            "password": "",
            "keyfile": "",
            "post_cmd": "",
            "memo": "",
            "group1": "",
            "group2": "",
            "group3": "",
        }
        data = extract_row_data(row)
        assert data["template"] == ""
