"""ttmacro.config の環境変数オーバーライドテスト。

config.py はモジュール import 時に環境変数を評価して定数を初期化するため、
テストは ``importlib.reload`` で再評価する。クラス終端の autouse フィクスチャで
環境クリーンな状態へリセットして他のテストファイルへの影響を防ぐ。
"""

from __future__ import annotations

import importlib
from collections.abc import Iterator
from pathlib import Path

import pytest

from ttmacro import config

ALL_ENV_VARS = (
    "TTMACRO_BASE_DIR",
    "TTMACRO_EXCEL_PATH",
    "TTMACRO_TEMPLATE_PATH",
    "TTMACRO_OUTPUT_DIR",
    "TTMACRO_LOGS_DIR",
    "TTMACRO_KEYS_DIR",
)


def _clear_env(monkeypatch: pytest.MonkeyPatch) -> None:
    """全 TTMACRO_* 環境変数をクリアする。"""
    for v in ALL_ENV_VARS:
        monkeypatch.delenv(v, raising=False)


class TestEnvOverride:
    """環境変数でパスを上書きできることのテスト。"""

    @pytest.fixture(autouse=True, scope="class")
    def _restore_defaults(self) -> Iterator[None]:
        """クラス内の全テスト終了後に config をデフォルトに戻す。"""
        yield
        importlib.reload(config)

    def test_default_paths_when_no_env(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """環境変数が一切ない場合、すべてリポジトリ既定値（BASE_DIR 起点）を使う。"""
        _clear_env(monkeypatch)
        importlib.reload(config)
        assert config.EXCEL_PATH == config.BASE_DIR / "data" / "servers.xlsx"
        assert config.TEMPLATE_PATH == config.BASE_DIR / "macros" / "template.ttl"
        assert config.OUTPUT_DIR == config.BASE_DIR / "macros"
        assert config.LOGS_DIR == config.BASE_DIR / "logs"
        assert config.KEYS_DIR == config.BASE_DIR / "keys"

    def test_base_dir_override_cascades(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """TTMACRO_BASE_DIR を変えると個別未指定の他パスもそれ起点になる。"""
        _clear_env(monkeypatch)
        monkeypatch.setenv("TTMACRO_BASE_DIR", str(tmp_path))
        importlib.reload(config)
        assert tmp_path.resolve() == config.BASE_DIR
        assert tmp_path.resolve() / "data" / "servers.xlsx" == config.EXCEL_PATH
        assert tmp_path.resolve() / "macros" == config.OUTPUT_DIR
        assert tmp_path.resolve() / "logs" == config.LOGS_DIR
        assert tmp_path.resolve() / "keys" == config.KEYS_DIR

    def test_individual_path_override(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """個別の環境変数で 1 つだけ上書きしても、他のパスは既定のまま。"""
        _clear_env(monkeypatch)
        custom = tmp_path / "custom" / "data" / "alt.xlsx"
        monkeypatch.setenv("TTMACRO_EXCEL_PATH", str(custom))
        importlib.reload(config)
        assert custom.resolve() == config.EXCEL_PATH
        # 他はデフォルト（BASE_DIR 起点）
        assert config.OUTPUT_DIR == config.BASE_DIR / "macros"
        assert config.TEMPLATE_PATH == config.BASE_DIR / "macros" / "template.ttl"

    def test_individual_overrides_take_precedence_over_base_dir(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        """個別の環境変数は TTMACRO_BASE_DIR より優先される。"""
        _clear_env(monkeypatch)
        base = tmp_path / "base"
        explicit_output = tmp_path / "elsewhere" / "macros"
        monkeypatch.setenv("TTMACRO_BASE_DIR", str(base))
        monkeypatch.setenv("TTMACRO_OUTPUT_DIR", str(explicit_output))
        importlib.reload(config)
        assert base.resolve() == config.BASE_DIR
        # OUTPUT_DIR は明示値
        assert explicit_output.resolve() == config.OUTPUT_DIR
        # 明示しなかった EXCEL_PATH は BASE_DIR 起点に追従
        assert base.resolve() / "data" / "servers.xlsx" == config.EXCEL_PATH

    def test_empty_env_var_treated_as_unset(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """空文字の環境変数は未設定として扱う（デフォルトを使う）。"""
        _clear_env(monkeypatch)
        monkeypatch.setenv("TTMACRO_EXCEL_PATH", "")
        importlib.reload(config)
        assert config.EXCEL_PATH == config.BASE_DIR / "data" / "servers.xlsx"


class TestEnvPath:
    """``_env_path`` ヘルパー関数のテスト。"""

    def test_returns_default_when_unset(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        monkeypatch.delenv("TTMACRO_TEST_VAR", raising=False)
        default = tmp_path / "default"
        assert config._env_path("TTMACRO_TEST_VAR", default) == default

    def test_returns_resolved_path_when_set(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        custom = tmp_path / "custom"
        monkeypatch.setenv("TTMACRO_TEST_VAR", str(custom))
        result = config._env_path("TTMACRO_TEST_VAR", tmp_path / "default")
        assert result == custom.resolve()

    def test_returns_default_when_empty_string(
        self, monkeypatch: pytest.MonkeyPatch, tmp_path: Path
    ) -> None:
        monkeypatch.setenv("TTMACRO_TEST_VAR", "")
        default = tmp_path / "default"
        assert config._env_path("TTMACRO_TEST_VAR", default) == default
