"""ttmacro.path_resolver のテスト。"""

from __future__ import annotations

from pathlib import Path

import pytest

from ttmacro import path_resolver


class TestCalculateRelativePath:
    """calculate_relative_path のテスト。"""

    def test_root_returns_empty_string(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """BASE_DIR 直下なら空文字を返す。"""
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        assert path_resolver.calculate_relative_path(tmp_path) == ""

    def test_one_level_deep_returns_one_dotdot(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """1階層深いなら '../' を返す。"""
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        assert path_resolver.calculate_relative_path(tmp_path / "macros") == "../"

    def test_three_levels_deep_returns_three_dotdots(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """3階層深いなら '../../../' を返す。"""
        monkeypatch.setattr("ttmacro.path_resolver.BASE_DIR", tmp_path)
        deep = tmp_path / "macros" / "g1" / "g2"
        assert path_resolver.calculate_relative_path(deep) == "../../../"


class TestGetLogDir:
    """get_log_dir のテスト。"""

    def test_output_root_returns_logs_root(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """OUTPUT_DIR 直下なら LOGS_DIR を返す。"""
        output_dir = tmp_path / "macros"
        logs_dir = tmp_path / "logs"
        monkeypatch.setattr("ttmacro.path_resolver.OUTPUT_DIR", output_dir)
        monkeypatch.setattr("ttmacro.path_resolver.LOGS_DIR", logs_dir)

        assert path_resolver.get_log_dir(output_dir) == logs_dir

    def test_subdir_returns_logs_with_same_subdir(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """OUTPUT_DIR/foo/bar なら LOGS_DIR/foo/bar を返す。"""
        output_dir = tmp_path / "macros"
        logs_dir = tmp_path / "logs"
        monkeypatch.setattr("ttmacro.path_resolver.OUTPUT_DIR", output_dir)
        monkeypatch.setattr("ttmacro.path_resolver.LOGS_DIR", logs_dir)

        target = output_dir / "home" / "prod"
        assert path_resolver.get_log_dir(target) == logs_dir / "home" / "prod"


class TestGetTargetDirectory:
    """get_target_directory のテスト。

    OUTPUT_DIR を tmp_path に差し替えて、実際にディレクトリを作成する挙動も検証する。
    """

    @pytest.fixture
    def patched_output_dir(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> Path:
        """OUTPUT_DIR を tmp_path に差し替える共通フィクスチャ。"""
        monkeypatch.setattr("ttmacro.path_resolver.OUTPUT_DIR", tmp_path)
        return tmp_path

    def test_no_group_returns_output_dir(self, patched_output_dir: Path) -> None:
        """group1 が空なら OUTPUT_DIR をそのまま返す（mkdir もしない）。"""
        data = {"group1": "", "group2": "", "group3": ""}
        assert path_resolver.get_target_directory(data) == patched_output_dir

    def test_group1_only(self, patched_output_dir: Path) -> None:
        """group1 のみ指定で1階層作成。"""
        data = {"group1": "envA", "group2": "", "group3": ""}
        result = path_resolver.get_target_directory(data)
        assert result == patched_output_dir / "envA"
        assert result.exists()
        assert result.is_dir()

    def test_full_three_levels(self, patched_output_dir: Path) -> None:
        """group1/2/3 すべて指定で3階層作成。"""
        data = {"group1": "envA", "group2": "type1", "group3": "subA"}
        result = path_resolver.get_target_directory(data)
        assert result == patched_output_dir / "envA" / "type1" / "subA"
        assert result.exists()

    def test_group2_without_group1_is_ignored(self, patched_output_dir: Path) -> None:
        """group1 が空なら子グループは無効（OUTPUT_DIR 直下に出力）。"""
        data = {"group1": "", "group2": "ignored", "group3": "ignored"}
        assert path_resolver.get_target_directory(data) == patched_output_dir

    def test_group3_without_group2_is_ignored(self, patched_output_dir: Path) -> None:
        """group2 が空なら group3 も無効。"""
        data = {"group1": "envA", "group2": "", "group3": "ignored"}
        result = path_resolver.get_target_directory(data)
        assert result == patched_output_dir / "envA"
