"""PyInstaller 用エントリスクリプト（GUI ランチャー）。

[project.scripts] の console shim を使わず、PyInstaller が直接食える単独
スクリプトとして ``ttmacro.launcher.main`` を呼ぶだけのアダプタ。
ビルド時に packaging/ttmacro-launcher.spec から参照される。
"""

from __future__ import annotations

from ttmacro.launcher import main

if __name__ == "__main__":
    main()
