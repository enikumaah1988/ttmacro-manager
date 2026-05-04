"""Tera Term TTL マクロランチャー（GUI）。

macros_root 配下の .ttl ファイルをツリー表示し、ダブルクリックや Enter で
Tera Term を起動するための Tkinter アプリケーション。
"""

from __future__ import annotations

import json
import logging
import re
import shutil
import subprocess
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any

import ttkbootstrap as ttk  # ttkbootstrap が ttk を拡張：API は ttk 互換 + bootstyle 引数

logger = logging.getLogger(__name__)

CONFIG_FILE = Path(__file__).resolve().parent / "launcher_config.json"
BASE_DIR = Path(__file__).resolve().parent.parent


def _parse_ttl_filename(filename: str) -> tuple[str, str, str]:
    """生成器のレイアウト '<name>_<host>_<user>.ttl' を逆解析する。

    Args:
        filename: TTL ファイル名（拡張子の有無は問わない）。

    Returns:
        ``(name, host, user)`` のタプル。レイアウトに合致しない場合
        （自作ファイル・アンダースコアが少ないファイル・host 部が
        IP/FQDN らしくないファイル等）は ``(拡張子なしのファイル名, '-', '-')``
        を返す。
    """
    stem = filename[:-4] if filename.endswith(".ttl") else filename
    # 後ろから 2 回分割して最大 3 要素にする（name に _ が含まれていてもよい）
    parts = stem.rsplit("_", 2)
    if len(parts) == 3:
        name, host, user = parts
        # host が IP または FQDN らしさを満たす
        # （英数字/ドット/ハイフンのみ、かつドットを含む or 先頭が数字）
        if re.match(r"^[a-zA-Z0-9.-]+$", host) and ("." in host or host[:1].isdigit()):
            return name, host, user
    return stem, "-", "-"


class LauncherApp:
    """TTL マクロランチャーの GUI アプリケーション。

    Tk ルートを受け取り、設定読み書き・ツリー構築・TTL 起動を司る。
    """

    def __init__(self, root: tk.Tk) -> None:
        """ウィンドウとウィジェットを構築する。

        Args:
            root: Tk のルートウィンドウ。
        """
        self.root = root
        self.root.title("Tera Term マクロランチャー")
        # 初期ウィンドウサイズ（ツリーで多めの行が見えるよう縦長め）
        self.root.geometry("1100x700")
        self.root.minsize(800, 500)

        self.macros_dir = tk.StringVar(master=root, value="")
        self.tterm_path = tk.StringVar(master=root, value="")
        self.search_query = tk.StringVar(master=root, value="")

        # notepad のフルパスを取得（Windows 標準なので通常見つかるが、見つからない場合は None）
        self.preferred_editor: str | None = shutil.which("notepad")

        config = self._load_config()
        self.tterm_path.set(config.get("teraterm_path", ""))
        self.macros_dir.set(config.get("macros_root", ""))

        self._build_config_frame()
        self._build_search_frame()
        self._build_tree_frame()
        self._build_bottom_frame()
        self._build_tree()

    # --- 設定の読み書き ---

    def _load_config(self) -> dict[str, str]:
        """launcher_config.json を読み込む。

        Returns:
            設定辞書。ファイル不在・読み込み失敗時は空辞書。
        """
        if not CONFIG_FILE.exists():
            return {}
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as e:
            logger.error("設定読み込みに失敗: %s", e)
            return {}

    def _save_config(self) -> None:
        """現在の設定値を launcher_config.json に書き込む。"""
        data = {
            "teraterm_path": self.tterm_path.get(),
            "macros_root": self.macros_dir.get(),
        }
        try:
            CONFIG_FILE.write_text(
                json.dumps(data, indent=4, ensure_ascii=False),
                encoding="utf-8",
            )
        except OSError as e:
            logger.error("設定保存に失敗: %s", e)
            messagebox.showerror("保存失敗", f"設定の保存に失敗しました:\n{e}")
            return
        messagebox.showinfo("保存完了", "設定を保存しました。")

    # --- TTL の起動・編集 ---

    def _get_selected_ttl_path(self) -> Path | None:
        """ツリーで選択されている TTL ファイルの絶対パスを返す。

        Returns:
            選択中 TTL の絶対パス。未選択・フォルダ選択時・取得失敗時は None。
        """
        selected = self.tree.selection()
        if not selected:
            return None

        # 'path' 列の値を取得（フォルダ行は空文字なので None 扱い）
        rel_path_str = self.tree.set(selected[0], "path")
        if not rel_path_str:
            return None

        try:
            ttl_path = Path(self.macros_dir.get()) / rel_path_str
            return ttl_path.resolve()
        except (OSError, ValueError) as e:
            logger.error("選択 TTL のパス解決に失敗: %s", e)
            return None

    def _run_ttl(self, ttl_path: Path) -> None:
        """Tera Term を起動して TTL を実行する。

        Args:
            ttl_path: 実行する TTL ファイルの絶対パス。
        """
        tterm = self.tterm_path.get()
        if not tterm or not Path(tterm).exists():
            messagebox.showerror(
                "エラー",
                f"Tera Term 実行ファイルが見つかりません:\n{tterm}",
            )
            return
        try:
            subprocess.Popen([tterm, f"/M={ttl_path}"])
        except OSError as e:
            messagebox.showerror("起動失敗", str(e))

    def _open_in_editor(self, ttl_path: Path) -> None:
        """TTL ファイルをエディタで開く。

        Args:
            ttl_path: 開く TTL ファイルの絶対パス。
        """
        # preferred_editor が None でも notepad を直接呼んで救済する
        editor = self.preferred_editor or "notepad"
        try:
            subprocess.Popen([editor, str(ttl_path)])
        except OSError as e:
            messagebox.showerror("起動失敗", f"エディタ起動に失敗しました:\n{e}")

    # --- イベントハンドラ ---

    def _on_double_click(self, event: object | None = None) -> None:
        """ツリー要素ダブルクリック時：選択 TTL を実行する。"""
        ttl_path = self._get_selected_ttl_path()
        if ttl_path:
            self._run_ttl(ttl_path)

    def _on_right_click(self, event: object | None = None) -> None:
        """右クリック時：選択 TTL をエディタで開く。"""
        ttl_path = self._get_selected_ttl_path()
        if ttl_path:
            self._open_in_editor(ttl_path)

    def _on_edit_button(self) -> None:
        """編集ボタン押下時：選択 TTL をエディタで開く。"""
        ttl_path = self._get_selected_ttl_path()
        if ttl_path:
            self._open_in_editor(ttl_path)

    # --- ツリー構築 ---

    def _build_tree(self) -> None:
        """macros_dir を再走査してツリーを再構築する。

        ``self.search_query`` が非空ならファイル名・グループパスに対する
        部分一致（大文字小文字無視）でフィルタする。フィルタ時は親フォルダを
        自動展開して見つけやすくする。template.ttl は常に対象外。
        """
        self.tree.delete(*self.tree.get_children())

        macros_root = Path(self.macros_dir.get())
        if not macros_root.exists():
            return

        filter_text = self.search_query.get().strip().lower()

        ungrouped: list[Path] = []

        for ttl_path in macros_root.rglob("*.ttl"):
            rel_parts = ttl_path.relative_to(macros_root).parts

            # macros/templates/ 配下はテンプレ置き場なのでツリー一覧から除外
            if rel_parts and rel_parts[0] == "templates":
                continue

            rel_path_str = str(ttl_path.relative_to(macros_root))
            # フィルタ：相対パス全体（フォルダ名 + ファイル名）に対する部分一致
            if filter_text and filter_text not in rel_path_str.lower():
                continue

            if len(rel_parts) == 1:
                ungrouped.append(ttl_path)
                continue

            parent = ""
            for i, part in enumerate(rel_parts):
                node_id = "/".join(rel_parts[: i + 1])
                if not self.tree.exists(node_id):
                    if i == len(rel_parts) - 1:
                        # leaf：ファイル名をパースして 4 列に分配
                        name, host, user = _parse_ttl_filename(ttl_path.name)
                        self.tree.insert(
                            parent,
                            "end",
                            iid=node_id,
                            text="",
                            values=[name, host, user, rel_path_str],
                        )
                    else:
                        self.tree.insert(parent, "end", iid=node_id, text=f"📁 {part}")
                parent = node_id

        if ungrouped:
            if not self.tree.exists("ungrouped"):
                self.tree.insert("", "end", iid="ungrouped", text="📁 未分類")
            for ttl_path in ungrouped:
                leaf_id = f"ungrouped/{ttl_path.name}"
                rel_path = str(ttl_path.relative_to(macros_root))
                name, host, user = _parse_ttl_filename(ttl_path.name)
                self.tree.insert(
                    "ungrouped",
                    "end",
                    iid=leaf_id,
                    text="",
                    values=[name, host, user, rel_path],
                )

        # フィルタ適用時は全フォルダを展開して見つけやすくする
        if filter_text:
            self._expand_all_folders()

    def _expand_all_folders(self) -> None:
        """ツリー上の全フォルダノードを展開する。"""
        for iid in self.tree.get_children():
            self._expand_recursive(iid)

    def _expand_recursive(self, iid: str) -> None:
        """再帰的にノードを展開する（子を持つノードのみ）。"""
        children = self.tree.get_children(iid)
        if children:
            self.tree.item(iid, open=True)
            for child in children:
                self._expand_recursive(child)

    def _on_clear_search(self) -> None:
        """検索ボックスをクリアしてツリーを再描画する。"""
        self.search_query.set("")
        self._build_tree()

    # --- ウィジェット構築 ---

    def _build_config_frame(self) -> None:
        """上部のパス入力フレームを構築する。

        ttkbootstrap のボタンは標準 ttk よりパディングが厚いため、
        grid セルに pady=4 を入れて行間が詰まって見えないようにする。
        """
        frame = ttk.Frame(self.root)
        frame.pack(fill=tk.X, padx=10, pady=5)

        # Tera Term 行
        ttk.Label(frame, text="Tera Termのパス:").grid(
            row=0, column=0, sticky="w", pady=4
        )
        ttk.Entry(frame, textvariable=self.tterm_path, width=60).grid(
            row=0, column=1, padx=5, pady=4
        )
        ttk.Button(frame, text="参照", command=self._on_browse_tterm, width=8).grid(
            row=0, column=2, padx=5, pady=4
        )
        ttk.Button(frame, text="設定保存", command=self._save_config, width=10).grid(
            row=0, column=3, padx=5, pady=4
        )

        # macros_root 行
        ttk.Label(frame, text="TTLマクロルート:").grid(
            row=1, column=0, sticky="w", pady=4
        )
        ttk.Entry(frame, textvariable=self.macros_dir, width=60).grid(
            row=1, column=1, padx=5, pady=4
        )
        ttk.Button(frame, text="参照", command=self._on_browse_macros, width=8).grid(
            row=1, column=2, padx=5, pady=4
        )
        ttk.Button(frame, text="ツリー更新", command=self._build_tree, width=10).grid(
            row=1, column=3, padx=5, pady=4
        )

    def _on_browse_tterm(self) -> None:
        """Tera Term 実行ファイルを参照ダイアログで選択する。

        キャンセル時（空文字が返る場合）は既存値を維持する。
        """
        path = filedialog.askopenfilename(filetypes=[("実行ファイル", "*.exe")])
        if path:
            self.tterm_path.set(path)

    def _on_browse_macros(self) -> None:
        """macros ルートディレクトリを参照ダイアログで選択する。

        キャンセル時（空文字が返る場合）は既存値を維持する。
        """
        path = filedialog.askdirectory()
        if path:
            self.macros_dir.set(path)

    def _build_search_frame(self) -> None:
        """設定フレームとツリーの間に検索バーを構築する。

        Entry の入力を即座にツリーに反映する（``<KeyRelease>`` で再構築）。
        """
        # 設定セクションと検索セクションの視覚的な境界線
        ttk.Separator(self.root, orient="horizontal").pack(
            fill=tk.X, padx=10, pady=(5, 5)
        )

        frame = ttk.Frame(self.root)
        frame.pack(fill=tk.X, padx=10, pady=(0, 5))

        ttk.Label(frame, text="検索:").pack(side=tk.LEFT)
        # ウィンドウ全幅まで伸ばさず固定幅。長い検索語が入る前提ではないため width=40
        entry = ttk.Entry(frame, textvariable=self.search_query, width=40)
        entry.pack(side=tk.LEFT, padx=5)
        # 入力ごとに即時反映（pandas/重い処理ではないので debounce 不要）
        entry.bind("<KeyRelease>", lambda _: self._build_tree())

        ttk.Button(frame, text="✕クリア", command=self._on_clear_search, width=8).pack(
            side=tk.LEFT
        )

    def _build_tree_frame(self) -> None:
        """中央のツリービューフレームを構築する。

        起動時には #0 / サーバ名 / IP / ユーザ名 が見切れずに表示されるよう
        固定幅とし、パス列だけ stretch=True で残りの幅を吸収させる。
        """
        frame = ttk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(
            frame,
            columns=("server", "ip", "user", "path"),
            show="tree headings",
        )
        self.tree.heading("#0", text="マクロ構成")
        self.tree.heading("server", text="サーバ名")
        self.tree.heading("ip", text="IP / ホスト")
        self.tree.heading("user", text="ユーザ名")
        self.tree.heading("path", text="TTL パス（相対）")

        self.tree.column("#0", anchor="w", width=200, stretch=False)
        # サーバ名は日本語や長めの名前を考慮して広めに
        self.tree.column("server", anchor="w", width=260, stretch=False)
        self.tree.column("ip", anchor="w", width=140, stretch=False)
        self.tree.column("user", anchor="w", width=100, stretch=False)
        # パス列は残幅を吸収（起動時にはみ出して隠れる場合あり、仕様）
        self.tree.column("path", anchor="w", width=400, stretch=True)

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Return>", self._on_double_click)
        self.tree.bind("<Button-3>", self._on_right_click)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

    def _build_bottom_frame(self) -> None:
        """下部のアクションボタンフレームを構築する。

        プライマリアクション（接続/編集/閉じる）は width=14 で広めに取り、
        最重要の「接続」のみ ttkbootstrap の success スタイルでハイライト。
        """
        frame = ttk.Frame(self.root)
        frame.pack(pady=10)

        # (label, command, bootstyle) の3要素
        actions: list[tuple[str, Callable[..., Any], str]] = [
            ("接続", self._on_double_click, "success"),
            ("編集", self._on_edit_button, "info"),
            ("閉じる", self.root.quit, "secondary"),
        ]
        for i, (label, cmd, style) in enumerate(actions):
            ttk.Button(frame, text=label, command=cmd, width=14, bootstyle=style).grid(
                row=0, column=i, padx=20
            )

    def run(self) -> None:
        """Tk のメインループを開始する。"""
        self.root.mainloop()


def main() -> None:
    """ランチャーアプリのエントリポイント。"""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    # ttkbootstrap の Window はテーマ付きの tk.Tk サブクラス
    root = ttk.Window(themename="darkly")
    app = LauncherApp(root)
    app.run()


if __name__ == "__main__":
    main()
