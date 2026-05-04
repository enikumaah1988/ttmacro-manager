"""ttmacro.cleaner のテスト。"""

from __future__ import annotations

from pathlib import Path

from ttmacro.cleaner import (
    delete_empty_subdirs,
    delete_ttl_files,
    find_empty_subdirs,
    find_ttl_files_to_delete,
)


class TestFindTtlFilesToDelete:
    """find_ttl_files_to_delete のテスト。"""

    def test_empty_dir(self, tmp_path: Path) -> None:
        assert find_ttl_files_to_delete(tmp_path) == []

    def test_nonexistent_dir(self, tmp_path: Path) -> None:
        assert find_ttl_files_to_delete(tmp_path / "nonexistent") == []

    def test_finds_ttl_at_multiple_depths(self, tmp_path: Path) -> None:
        (tmp_path / "a.ttl").touch()
        (tmp_path / "sub").mkdir()
        (tmp_path / "sub" / "b.ttl").touch()
        (tmp_path / "sub" / "deep").mkdir()
        (tmp_path / "sub" / "deep" / "c.ttl").touch()
        result = find_ttl_files_to_delete(tmp_path)
        assert len(result) == 3
        assert {p.name for p in result} == {"a.ttl", "b.ttl", "c.ttl"}

    def test_excludes_template_ttl(self, tmp_path: Path) -> None:
        (tmp_path / "a.ttl").touch()
        (tmp_path / "template.ttl").touch()
        result = find_ttl_files_to_delete(tmp_path)
        assert len(result) == 1
        assert result[0].name == "a.ttl"

    def test_template_ttl_in_subdirectory_also_excluded(self, tmp_path: Path) -> None:
        """サブディレクトリ内の template.ttl も除外される（厳密にはあり得ない構成だが、防御）。"""
        (tmp_path / "sub").mkdir()
        (tmp_path / "sub" / "template.ttl").touch()
        (tmp_path / "sub" / "real.ttl").touch()
        result = find_ttl_files_to_delete(tmp_path)
        assert {p.name for p in result} == {"real.ttl"}

    def test_ignores_other_extensions(self, tmp_path: Path) -> None:
        (tmp_path / "a.ttl").touch()
        (tmp_path / "b.txt").touch()
        (tmp_path / "c.log").touch()
        result = find_ttl_files_to_delete(tmp_path)
        assert len(result) == 1
        assert result[0].name == "a.ttl"


class TestFindEmptySubdirs:
    """find_empty_subdirs のテスト。"""

    def test_no_subdirs(self, tmp_path: Path) -> None:
        assert find_empty_subdirs(tmp_path) == []

    def test_single_empty_subdir(self, tmp_path: Path) -> None:
        (tmp_path / "empty").mkdir()
        result = find_empty_subdirs(tmp_path)
        assert len(result) == 1
        assert result[0].name == "empty"

    def test_excludes_output_dir_itself(self, tmp_path: Path) -> None:
        # output_dir そのものは結果に含めない
        result = find_empty_subdirs(tmp_path)
        assert tmp_path not in result

    def test_excludes_dirs_with_content(self, tmp_path: Path) -> None:
        d = tmp_path / "with_content"
        d.mkdir()
        (d / "file.txt").touch()
        assert find_empty_subdirs(tmp_path) == []

    def test_only_truly_empty_dirs(self, tmp_path: Path) -> None:
        """中間ディレクトリは子を持つので空ではない（一回の呼び出しでは末端のみ返す）。"""
        deep = tmp_path / "a" / "b" / "c"
        deep.mkdir(parents=True)
        # a, a/b は子ディレクトリを持つので非空。c のみ空
        result = find_empty_subdirs(tmp_path)
        assert len(result) == 1
        assert result[0].name == "c"

    def test_depth_descending_order(self, tmp_path: Path) -> None:
        """深いディレクトリが先に返る。"""
        # 浅い空: tmp_path/shallow
        # 深い空: tmp_path/container/deep
        (tmp_path / "shallow").mkdir()
        (tmp_path / "container" / "deep").mkdir(parents=True)
        result = find_empty_subdirs(tmp_path)
        names = [p.name for p in result]
        # 'deep' (3階層) が 'shallow' (2階層) より先
        assert names == ["deep", "shallow"]


class TestDeleteTtlFiles:
    """delete_ttl_files のテスト。"""

    def test_deletes_files(self, tmp_path: Path) -> None:
        f1 = tmp_path / "a.ttl"
        f2 = tmp_path / "b.ttl"
        f1.touch()
        f2.touch()
        count = delete_ttl_files([f1, f2])
        assert count == 2
        assert not f1.exists()
        assert not f2.exists()

    def test_returns_zero_on_empty_list(self, tmp_path: Path) -> None:
        assert delete_ttl_files([]) == 0

    def test_continues_on_individual_failure(self, tmp_path: Path) -> None:
        """1 件失敗しても残りを処理する。"""
        f1 = tmp_path / "a.ttl"
        f1.touch()
        f_missing = tmp_path / "missing.ttl"  # 存在しない
        count = delete_ttl_files([f1, f_missing])
        assert count == 1
        assert not f1.exists()


class TestDeleteEmptySubdirs:
    """delete_empty_subdirs のテスト。"""

    def test_removes_sibling_empty_dirs(self, tmp_path: Path) -> None:
        (tmp_path / "empty1").mkdir()
        (tmp_path / "empty2").mkdir()
        count = delete_empty_subdirs(tmp_path)
        assert count == 2
        assert not (tmp_path / "empty1").exists()
        assert not (tmp_path / "empty2").exists()

    def test_cascading_removal(self, tmp_path: Path) -> None:
        """tmp_path/a/b/c (全部空) が連鎖的に削除される。"""
        deep = tmp_path / "a" / "b" / "c"
        deep.mkdir(parents=True)
        count = delete_empty_subdirs(tmp_path)
        assert count == 3
        assert not (tmp_path / "a").exists()

    def test_does_not_remove_output_dir(self, tmp_path: Path) -> None:
        """output_dir 自体が空でも消されない。"""
        delete_empty_subdirs(tmp_path)
        assert tmp_path.exists()

    def test_does_not_remove_nonempty_dirs(self, tmp_path: Path) -> None:
        d = tmp_path / "with_file"
        d.mkdir()
        (d / "f.txt").touch()
        count = delete_empty_subdirs(tmp_path)
        assert count == 0
        assert d.exists()

    def test_partial_cascade(self, tmp_path: Path) -> None:
        """a/b/c (空) と a/d/file.txt (中身あり) が混在する場合、
        c, b は削除されるが d は中身があるので残り、a も残る。"""
        (tmp_path / "a" / "b" / "c").mkdir(parents=True)
        (tmp_path / "a" / "d").mkdir(parents=True)
        (tmp_path / "a" / "d" / "f.txt").touch()
        count = delete_empty_subdirs(tmp_path)
        assert count == 2  # c と b
        assert not (tmp_path / "a" / "b").exists()
        assert (tmp_path / "a" / "d").exists()
        assert (tmp_path / "a").exists()

    def test_returns_zero_on_empty_dir(self, tmp_path: Path) -> None:
        # 空の output_dir では削除対象なし
        assert delete_empty_subdirs(tmp_path) == 0
