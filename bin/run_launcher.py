"""Tera Term TTL マクロランチャー（GUI）。

macros_root 配下の .ttl ファイルをツリー表示し、ダブルクリックや Enter で
Tera Term を起動するための Tkinter アプリケーション。
"""

from __future__ import annotations

import json
import logging
import shutil
import subprocess
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

logger = logging.getLogger(__name__)

CONFIG_FILE = Path(__file__).resolve().parent / "launcher_config.json"
BASE_DIR = Path(__file__).resolve().parent.parent


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

        self.macros_dir = tk.StringVar(master=root, value="")
        self.tterm_path = tk.StringVar(master=root, value="")

        # notepad のフルパスを取得（Windows 標準なので通常見つかるが、見つからない場合は None）
        self.preferred_editor: str | None = shutil.which("notepad")

        config = self._load_config()
        self.tterm_path.set(config.get("teraterm_path", ""))
        self.macros_dir.set(config.get("macros_root", ""))

        self._build_config_frame()
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
            選択中 TTL の絶対パス。未選択・取得失敗時は None。
        """
        selected = self.tree.selection()
        if not selected:
            return None

        values = self.tree.item(selected[0], "values")
        if not values:
            return None

        try:
            relative_path = Path(values[0])
            ttl_path = Path(self.macros_dir.get()) / relative_path
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

        サブディレクトリ階層をフォルダノードとして表示し、直下の TTL は
        「未分類」フォルダに集める。template.ttl は対象外。
        """
        self.tree.delete(*self.tree.get_children())

        macros_root = Path(self.macros_dir.get())
        if not macros_root.exists():
            return

        ungrouped: list[Path] = []

        for ttl_path in macros_root.rglob("*.ttl"):
            if ttl_path.name == "template.ttl":
                continue
            rel_parts = ttl_path.relative_to(macros_root).parts
            if len(rel_parts) == 1:
                ungrouped.append(ttl_path)
                continue

            parent = ""
            for i, part in enumerate(rel_parts):
                node_id = "/".join(rel_parts[: i + 1])
                if not self.tree.exists(node_id):
                    if i == len(rel_parts) - 1:
                        rel_path = str(ttl_path.relative_to(macros_root))
                        self.tree.insert(
                            parent, "end", iid=node_id, text=part, values=[rel_path]
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
                self.tree.insert(
                    "ungrouped",
                    "end",
                    iid=leaf_id,
                    text=ttl_path.name,
                    values=[rel_path],
                )

    # --- ウィジェット構築 ---

    def _build_config_frame(self) -> None:
        """上部のパス入力フレームを構築する。"""
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.X, padx=10, pady=5)

        # Tera Term 行
        tk.Label(frame, text="Tera Termのパス:").grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.tterm_path, width=60).grid(
            row=0, column=1, padx=5
        )
        tk.Button(
            frame,
            text="参照",
            command=lambda: self.tterm_path.set(
                filedialog.askopenfilename(filetypes=[("実行ファイル", "*.exe")])
            ),
        ).grid(row=0, column=2, padx=5)
        tk.Button(frame, text="保存", command=self._save_config).grid(
            row=0, column=3, padx=5
        )

        # macros_root 行
        tk.Label(frame, text="TTLマクロルート:").grid(row=1, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.macros_dir, width=60).grid(
            row=1, column=1, padx=5
        )
        tk.Button(
            frame,
            text="参照",
            command=lambda: self.macros_dir.set(filedialog.askdirectory()),
        ).grid(row=1, column=2, padx=5)
        tk.Button(frame, text="再読込", command=self._build_tree).grid(
            row=1, column=3, padx=5
        )

    def _build_tree_frame(self) -> None:
        """中央のツリービューフレームを構築する。"""
        frame = tk.Frame(self.root)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.tree = ttk.Treeview(frame, columns=("path",), show="tree headings")
        self.tree.heading("#0", text="マクロ構成")
        self.tree.heading("path", text="TTLマクロ格納パス（相対）")
        self.tree.column("#0", anchor="w", width=300)
        self.tree.column("path", anchor="w", width=500)

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Return>", self._on_double_click)
        self.tree.bind("<Button-3>", self._on_right_click)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

    def _build_bottom_frame(self) -> None:
        """下部のアクションボタンフレームを構築する。"""
        frame = tk.Frame(self.root)
        frame.pack(pady=10)

        actions: list[tuple[str, object]] = [
            ("接続実行", self._on_double_click),
            ("編集", self._on_edit_button),
            ("閉じる", self.root.quit),
        ]
        for i, (label, cmd) in enumerate(actions):
            tk.Button(frame, text=label, command=cmd).grid(row=0, column=i, padx=20)

    def run(self) -> None:
        """Tk のメインループを開始する。"""
        self.root.mainloop()


def main() -> None:
    """ランチャーアプリのエントリポイント。"""
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    )
    root = tk.Tk()
    app = LauncherApp(root)
    app.run()


if __name__ == "__main__":
    main()
