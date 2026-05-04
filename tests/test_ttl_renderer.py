"""ttmacro.ttl_renderer のテスト。"""

from __future__ import annotations

from pathlib import Path

import pytest

from ttmacro.ttl_renderer import generate_ttl_content, sanitize_name


class TestSanitizeName:
    """sanitize_name の Windows 禁止文字置換テスト。"""

    @pytest.mark.parametrize(
        ("input_name", "expected"),
        [
            ("normal_name", "normal_name"),
            ("with space", "with space"),
            ("forbidden\\char", "forbidden_char"),
            ("multi/forbidden:chars", "multi_forbidden_chars"),
            ('all*?"<>|chars', "all______chars"),
            ("日本語の名前", "日本語の名前"),
            ("", ""),
            ("Path\\to\\file:name", "Path_to_file_name"),
            ("hyphen-and_underscore", "hyphen-and_underscore"),
        ],
    )
    def test_replaces_forbidden_chars(self, input_name: str, expected: str) -> None:
        assert sanitize_name(input_name) == expected


class TestGenerateTtlContent:
    """generate_ttl_content のプレースホルダ展開テスト。"""

    @pytest.fixture
    def base_data(self) -> dict[str, str]:
        return {
            "name": "infra01",
            "host": "192.168.0.10",
            "port": "22",
            "user": "rocky",
            "password": "secret",
            "keyfile_name": "id_ed25519.ppk",
            "post_cmd": "",
            "memo": "本社NAS",
            "group1": "LocationA",
            "group2": "",
            "group3": "",
        }

    def test_all_placeholders_substituted(
        self, base_data: dict[str, str], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """主要プレースホルダがすべて置換される。"""
        # calculate_relative_path が BASE_DIR 起点で計算するので tmp_path に合わせる
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        target_dir = tmp_path / "macros" / "LocationA"
        target_dir.mkdir(parents=True)

        template = (
            "name={name}\nhost={hostname}\nport={port}\nuser={username}\n"
            "pw={password}\nkey={keyfile}\nrel={rel_path}\n"
            "ts={created_at}\nmemo={memo}\npc={post_commands}\n"
        )
        result = generate_ttl_content(base_data, template, "2026/01/01 00:00:00", target_dir)

        assert "name=infra01" in result
        assert "host=192.168.0.10" in result
        assert "port=22" in result
        assert "user=rocky" in result
        assert "pw=secret" in result
        assert "key=id_ed25519.ppk" in result
        assert "rel=../../" in result  # macros/LocationA = 2 階層
        assert "ts=2026/01/01 00:00:00" in result
        assert "memo=本社NAS" in result
        assert "pc=" in result  # post_cmd 空なので post_commands も空

    def test_post_commands_single_line(
        self, base_data: dict[str, str], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """1行のポストコマンドが wait/sendln ペアに展開される。"""
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        base_data["post_cmd"] = "ls -la"

        result = generate_ttl_content(base_data, "{post_commands}", "ts", tmp_path)
        assert "wait '$' '#'" in result
        assert "sendln 'ls -la'" in result

    def test_post_commands_multiple_lines(
        self, base_data: dict[str, str], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """複数行のポストコマンドがそれぞれ展開される。"""
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        base_data["post_cmd"] = "ls -la\ndate\nwhoami"

        result = generate_ttl_content(base_data, "{post_commands}", "ts", tmp_path)
        assert "sendln 'ls -la'" in result
        assert "sendln 'date'" in result
        assert "sendln 'whoami'" in result

    def test_post_commands_empty_lines_ignored(
        self, base_data: dict[str, str], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """空行はポストコマンドから除外される。"""
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        base_data["post_cmd"] = "ls\n\n   \ndate"

        result = generate_ttl_content(base_data, "{post_commands}", "ts", tmp_path)
        # 空行が混じっていても sendln は 2 つだけ
        assert result.count("sendln") == 2
