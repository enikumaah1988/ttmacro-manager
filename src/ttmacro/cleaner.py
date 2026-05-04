"""生成済み TTL マクロと空ディレクトリの削除モジュール。

CLI の ``--clean`` オプション時に呼び出され、グループ階層変更で残った
旧 TTL ファイルを掃除する。``template.ttl`` は対象外。
"""

from __future__ import annotations

from pathlib import Path


def find_ttl_files_to_delete(output_dir: Path) -> list[Path]:
    """削除対象の TTL 一覧を返す（``template.ttl`` 以外の ``*.ttl`` 全件）。

    Args:
        output_dir: macros ルートディレクトリ。

    Returns:
        削除対象 TTL の絶対パス一覧。``output_dir`` 不在時は空リスト。
    """
    if not output_dir.exists():
        return []

    return [ttl for ttl in output_dir.rglob("*.ttl") if ttl.name != "template.ttl"]


def find_empty_subdirs(output_dir: Path) -> list[Path]:
    """``output_dir`` 配下の空ディレクトリ一覧を深度降順で返す。

    ``output_dir`` そのものは結果に含めない（最上位は保持）。

    Args:
        output_dir: 探索ルートディレクトリ。

    Returns:
        空ディレクトリのパス一覧（深いものが先）。
    """
    if not output_dir.exists():
        return []

    empty_dirs: list[Path] = []
    for subdir in output_dir.rglob("*"):
        if not subdir.is_dir() or subdir == output_dir:
            continue
        try:
            if not any(subdir.iterdir()):
                empty_dirs.append(subdir)
        except OSError:
            # 権限エラー等は無視
            continue

    # 深度降順（深いディレクトリを先に rmdir すれば連鎖的に親も空になる）
    empty_dirs.sort(key=lambda p: len(p.parts), reverse=True)
    return empty_dirs


def delete_ttl_files(files: list[Path]) -> int:
    """指定された TTL ファイルを削除する。

    個別の削除失敗（権限・既に削除済み等）は無視して継続する。

    Args:
        files: 削除対象のファイル一覧。

    Returns:
        実際に削除に成功した件数。
    """
    count = 0
    for ttl in files:
        try:
            ttl.unlink()
            count += 1
        except OSError:
            pass
    return count


def delete_empty_subdirs(output_dir: Path) -> int:
    """``output_dir`` 配下の空ディレクトリを連鎖的に削除する。

    一度の削除で親ディレクトリが新たに空になるケースに対応するため、
    空ディレクトリが見つからなくなるまで繰り返す。``output_dir`` 自体は
    削除しない。

    Args:
        output_dir: 探索ルートディレクトリ。

    Returns:
        削除に成功したディレクトリ数の合計。
    """
    total = 0
    while True:
        empty_dirs = find_empty_subdirs(output_dir)
        if not empty_dirs:
            break
        progressed = False
        for d in empty_dirs:
            try:
                d.rmdir()
                total += 1
                progressed = True
            except OSError:
                # 削除失敗（権限等）はスキップ
                pass
        # 削除がひとつも進まなかった場合は無限ループ防止のため抜ける
        if not progressed:
            break
    return total
