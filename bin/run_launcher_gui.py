import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import configparser
import subprocess

# GUI構築の前にrootを定義
root = tk.Tk()
root.title("Tera Term マクロランチャー")

# プロジェクトルートの取得
BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE_DIR / "bin" / "config.ini"
MACROS_DIR = tk.StringVar(master=root, value=str(BASE_DIR / "macros"))
TTERM_PATH = tk.StringVar(master=root, value="")

VERSION = "v1.0.0"

# 設定ファイル読み込みと保存
def load_config():
    config = configparser.ConfigParser()
    config.read(CONFIG_PATH, encoding="utf-8")
    if config.has_option("launcher", "teraterm_path"):
        TTERM_PATH.set(config.get("launcher", "teraterm_path"))

def save_config():
    config = configparser.ConfigParser()
    config["launcher"] = {"teraterm_path": TTERM_PATH.get()}
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        config.write(f)

# TTL実行処理
def run_ttl(ttl_path, tterm_path):
    if not tterm_path or not Path(tterm_path).exists():
        messagebox.showerror("エラー", f"Tera Term 実行ファイルが見つかりません:\n{tterm_path}")
        return
    try:
        fullpath = ttl_path.resolve()
        subprocess.Popen([str(tterm_path), f'/M={str(fullpath)}'])
    except Exception as e:
        messagebox.showerror("起動失敗", str(e))

# TTLファイルをツリー構造でリストアップ
def build_tree(tree):
    tree.delete(*tree.get_children())
    grouped_nodes = {}
    ungrouped_ttls = []
    macro_root = Path(MACROS_DIR.get())

    for ttl_path in macro_root.rglob("*.ttl"):
        if ttl_path.name == "template.ttl":
            continue
        rel_parts = ttl_path.relative_to(macro_root).parts
        if len(rel_parts) == 1:
            ungrouped_ttls.append(ttl_path)
            continue

        parent = ''
        for i, part in enumerate(rel_parts):
            node_id = '/'.join(rel_parts[:i+1])
            if not tree.exists(node_id):
                if i == len(rel_parts) - 1:
                    rel_path = str(ttl_path.relative_to(BASE_DIR))
                    tree.insert(parent, "end", iid=node_id, text=part, values=[rel_path])
                else:
                    tree.insert(parent, "end", iid=node_id, text=f"📁 {part}")
            parent = node_id

    if ungrouped_ttls:
        if not tree.exists('ungrouped'):
            tree.insert('', 'end', iid='ungrouped', text='📁 未分類')
        for ttl_path in ungrouped_ttls:
            leaf_id = f"ungrouped/{ttl_path.name}"
            rel_path = str(ttl_path.relative_to(BASE_DIR))
            tree.insert('ungrouped', "end", iid=leaf_id, text=ttl_path.name, values=[rel_path])

# 実行処理
def on_double_click(event):
    selected = tree.selection()
    if not selected:
        return
    path = tree.item(selected[0], 'values')
    if path:
        ttl_absolute_path = BASE_DIR / path[0]
        run_ttl(ttl_absolute_path, TTERM_PATH.get())

# 初期化
load_config()

# Tera TermパスとTTLルートの入力欄
frame_config = tk.Frame(root)
frame_config.pack(fill=tk.X, padx=10, pady=5)

tk.Label(frame_config, text="Tera Termのパス:").grid(row=0, column=0, sticky="w")
entry_tterm = tk.Entry(frame_config, textvariable=TTERM_PATH, width=60)
entry_tterm.grid(row=0, column=1, padx=5)

btn_browse_tterm = tk.Button(frame_config, text="参照", command=lambda: TTERM_PATH.set(filedialog.askopenfilename(filetypes=[("実行ファイル", "*.exe")])) )
btn_browse_tterm.grid(row=0, column=2, padx=5)

btn_save = tk.Button(frame_config, text="保存", command=save_config)
btn_save.grid(row=0, column=3, padx=5)

tk.Label(frame_config, text="TTLマクロルート:").grid(row=1, column=0, sticky="w")
entry_macros = tk.Entry(frame_config, textvariable=MACROS_DIR, width=60)
entry_macros.grid(row=1, column=1, padx=5)

btn_browse_macros = tk.Button(frame_config, text="参照", command=lambda: MACROS_DIR.set(filedialog.askdirectory()))
btn_browse_macros.grid(row=1, column=2, padx=5)

btn_reload = tk.Button(frame_config, text="再読込", command=lambda: build_tree(tree))
btn_reload.grid(row=1, column=3, padx=5)

# TTL一覧表示
frame_tree = tk.Frame(root)
frame_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

tree = ttk.Treeview(frame_tree, columns=("path",), show="tree headings")
tree.heading("#0", text="マクロ構成")
tree.heading("path", text="TTLマクロ格納パス（相対）")
tree.column("#0", anchor="w", width=300)
tree.column("path", anchor="w", width=500)
tree.bind("<Double-1>", on_double_click)

tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(frame_tree, orient="vertical", command=tree.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
tree.configure(yscrollcommand=scrollbar.set)

# 実行・終了ボタン（中央寄せ）
frame_bottom = tk.Frame(root)
frame_bottom.pack(pady=10)

btn_run = tk.Button(frame_bottom, text="接続実行", command=lambda: on_double_click(None))
btn_run.grid(row=0, column=0, padx=20)
btn_close = tk.Button(frame_bottom, text="閉じる", command=root.quit)
btn_close.grid(row=0, column=1, padx=20)

build_tree(tree)
root.mainloop()
