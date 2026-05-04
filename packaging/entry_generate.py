"""PyInstaller 用エントリスクリプト（TTL 生成 CLI）。

[project.scripts] の console shim を使わず、PyInstaller が直接食える単独
スクリプトとして ``ttmacro.cli.main`` を呼ぶだけのアダプタ。
ビルド時に packaging/ttmacro-generate.spec から参照される。
"""

from __future__ import annotations

from ttmacro.cli import main

if __name__ == "__main__":
    main()
