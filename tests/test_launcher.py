"""ttmacro.launcher の純粋関数テスト。

GUI 部分（``LauncherApp``）は Tk ループを起こすためテスト対象外。
ファイル名パーサ ``_parse_ttl_filename`` のみを対象とする。
"""

from __future__ import annotations

import pytest

from ttmacro.launcher import _parse_ttl_filename


class TestParseTtlFilename:
    """_parse_ttl_filename のテスト。

    レイアウトは生成器の規約 ``<name>_<host>_<user>.ttl`` を逆解析する。
    host 部が IP か FQDN らしい時のみ 3 列を返し、そうでなければ
    ``(stem, '-', '-')`` で表示する仕様。
    """

    @pytest.mark.parametrize(
        "filename,expected",
        [
            # 標準的な IP ホスト
            ("infra01_192.168.0.10_rocky.ttl", ("infra01", "192.168.0.10", "rocky")),
            # FQDN（ドット含む）
            (
                "web01_server.example.com_admin.ttl",
                ("web01", "server.example.com", "admin"),
            ),
            # name にハイフンを含む
            ("web-01_192.168.1.5_root.ttl", ("web-01", "192.168.1.5", "root")),
            # name にアンダースコアを含む（rsplit で後ろから 2 回分割）
            ("my_server_01_10.0.0.1_admin.ttl", ("my_server_01", "10.0.0.1", "admin")),
            # 拡張子なしでも動く
            ("srv_192.168.0.1_user", ("srv", "192.168.0.1", "user")),
        ],
    )
    def test_valid_layouts_parse_correctly(
        self, filename: str, expected: tuple[str, str, str]
    ) -> None:
        assert _parse_ttl_filename(filename) == expected

    def test_underscore_count_too_few_returns_fallback(self) -> None:
        """アンダースコアが 1 つ以下なら 3 分割できないので fallback 表示。"""
        assert _parse_ttl_filename("noseparator.ttl") == ("noseparator", "-", "-")
        assert _parse_ttl_filename("only_one.ttl") == ("only_one", "-", "-")

    def test_host_part_not_ip_or_fqdn_returns_fallback(self) -> None:
        """host 部が IP/FQDN らしくない（ドット無し+先頭非数字）と fallback。"""
        # 'web' は単一トークンで FQDN でも IP でもない
        assert _parse_ttl_filename("a_web_user.ttl") == ("a_web_user", "-", "-")

    def test_host_part_with_invalid_chars_returns_fallback(self) -> None:
        """host 部に英数字・ドット・ハイフン以外が入ると弾かれる。"""
        # スペース等が含まれるケース（実際には起きにくいが防御的に）
        assert _parse_ttl_filename("a_bad host_u.ttl") == ("a_bad host_u", "-", "-")

    def test_extension_only_ttl_is_stripped(self) -> None:
        """末尾 .ttl のみ除去される（他拡張子は user 部に残る仕様）。"""
        # 現状仕様: user 部は検証していないので拡張子が混入しても通る
        assert _parse_ttl_filename("srv_192.168.0.1_user.txt") == (
            "srv",
            "192.168.0.1",
            "user.txt",
        )
