"""ttmacro.ttl_renderer のテスト。"""

from __future__ import annotations

from pathlib import Path

import pytest

from ttmacro.ttl_renderer import (
    generate_ttl_content,
    load_template,
    resolve_template_path,
    sanitize_name,
)


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
        result = generate_ttl_content(
            base_data, template, "2026/01/01 00:00:00", target_dir
        )

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

    def test_password_containing_placeholder_is_not_re_expanded(
        self, base_data: dict[str, str], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """password の値に ``{memo}`` が含まれていても再展開されない（単一パス置換）。"""
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        base_data["password"] = "abc{memo}def"
        base_data["memo"] = "LEAKED"

        result = generate_ttl_content(
            base_data, "pw={password}\nmemo={memo}", "ts", tmp_path
        )
        # password の中の ``{memo}`` は文字列として残り、memo の値で置換されない
        assert "pw=abc{memo}def" in result
        assert "memo=LEAKED" in result
        assert "abcLEAKEDdef" not in result

    def test_memo_containing_placeholder_kept_as_literal(
        self, base_data: dict[str, str], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """memo の値に ``{password}`` 文字列が含まれていても展開されず文字列として残る。"""
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        base_data["password"] = "secret"
        base_data["memo"] = "see {password} field"

        result = generate_ttl_content(
            base_data, "pw={password}\nmemo={memo}", "ts", tmp_path
        )
        assert "pw=secret" in result
        # memo に書いた ``{password}`` は文字列として残る（password 値に再展開されない）
        assert "memo=see {password} field" in result

    def test_unknown_placeholder_left_intact(
        self, base_data: dict[str, str], tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """登録されていないプレースホルダ ``{foo}`` はそのまま残す。"""
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        result = generate_ttl_content(base_data, "x={foo} y={name}", "ts", tmp_path)
        assert "x={foo}" in result  # 未知プレースホルダはそのまま
        assert "y=infra01" in result


class TestResolveTemplatePath:
    """resolve_template_path のテスト。"""

    def test_empty_returns_default(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """空文字ならデフォルト TEMPLATE_PATH を返す。"""
        default = tmp_path / "macros" / "template.ttl"
        monkeypatch.setattr("ttmacro.ttl_renderer.TEMPLATE_PATH", default)
        assert resolve_template_path("") == default

    def test_named_resolves_under_templates_dir(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """名前指定で TEMPLATES_DIR/<name>.ttl を返す。"""
        templates = tmp_path / "macros" / "templates"
        templates.mkdir(parents=True)
        monkeypatch.setattr("ttmacro.ttl_renderer.TEMPLATES_DIR", templates)
        result = resolve_template_path("nodejs")
        assert result == (templates / "nodejs.ttl").resolve()

    def test_extension_optional(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """``.ttl`` 拡張子付きで指定しても付けなくても同じ結果。"""
        templates = tmp_path / "macros" / "templates"
        templates.mkdir(parents=True)
        monkeypatch.setattr("ttmacro.ttl_renderer.TEMPLATES_DIR", templates)
        with_ext = resolve_template_path("nodejs.ttl")
        without_ext = resolve_template_path("nodejs")
        assert with_ext == without_ext

    def test_subpath_allowed(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """サブディレクトリを含む名前も許容（'subdir/foo' → templates/subdir/foo.ttl）。"""
        templates = tmp_path / "macros" / "templates"
        templates.mkdir(parents=True)
        monkeypatch.setattr("ttmacro.ttl_renderer.TEMPLATES_DIR", templates)
        result = resolve_template_path("subdir/foo")
        assert result == (templates / "subdir" / "foo.ttl").resolve()

    def test_path_traversal_rejected(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """``../`` で TEMPLATES_DIR の外を指すとエラー。"""
        templates = tmp_path / "macros" / "templates"
        templates.mkdir(parents=True)
        monkeypatch.setattr("ttmacro.ttl_renderer.TEMPLATES_DIR", templates)
        with pytest.raises(ValueError, match="TEMPLATES_DIR の外"):
            resolve_template_path("../../etc/passwd")


class TestLoadTemplateWithPath:
    """load_template の path 引数対応テスト。"""

    def test_loads_specified_path(self, tmp_path: Path) -> None:
        custom = tmp_path / "custom.ttl"
        custom.write_text("custom content", encoding="utf-8")
        assert load_template(custom) == "custom content"

    def test_default_when_no_path(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """path 省略時は TEMPLATE_PATH を読む（後方互換）。"""
        default = tmp_path / "default.ttl"
        default.write_text("default content", encoding="utf-8")
        monkeypatch.setattr("ttmacro.ttl_renderer.TEMPLATE_PATH", default)
        assert load_template() == "default content"

    def test_raises_when_specified_path_missing(self, tmp_path: Path) -> None:
        missing = tmp_path / "nonexistent.ttl"
        with pytest.raises(FileNotFoundError):
            load_template(missing)

    def test_raises_when_empty_file(self, tmp_path: Path) -> None:
        empty = tmp_path / "empty.ttl"
        empty.write_text("", encoding="utf-8")
        with pytest.raises(ValueError, match="空"):
            load_template(empty)
