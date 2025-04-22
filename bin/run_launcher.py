import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import subprocess
import shutil
import json

# --- 設定ファイルとエディタの定義 ---
CONFIG_FILE = Path(__file__).resolve().parent / "launcher_config.json"
BASE_DIR = Path(__file__).resolve().parent.parent
PREFERRED_EDITOR = shutil.which("notepad")

# --- GUI初期化 ---
root = tk.Tk()
root.title("Tera Term マクロランチャー")

# --- 設定値の初期化 ---
MACROS_DIR = tk.StringVar(master=root, value="")
TTERM_PATH = tk.StringVar(master=root, value="")

# --- 設定読み書き関数 ---
def load_launcher_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"[設定読み込み失敗] {e}")
    return {}

def save_launcher_config(teraterm_path: str, macros_root: str):
    data = {
        "teraterm_path": teraterm_path,
        "macros_root": macros_root
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        print(f"[設定保存失敗] {e}")

def save_config():
    save_launcher_config(TTERM_PATH.get(), MACROS_DIR.get())
    messagebox.showinfo("保存完了", "設定を保存しました。")

# --- TTLファイル起動・編集関係 ---
def run_ttl(ttl_path, tterm_path):
    if not tterm_path or not Path(tterm_path).exists():
        messagebox.showerror("エラー", f"Tera Term 実行ファイルが見つかりません:\n{tterm_path}")
        return
    try:
        subprocess.Popen([str(tterm_path), f'/M={str(ttl_path.resolve())}'])
    except Exception as e:
        messagebox.showerror("起動失敗", str(e))

def get_selected_ttl_path():
    selected = tree.selection()
    if not selected:
        return None

    values = tree.item(selected[0], "values")
    if not values:
        return None

    try:
        relative_path = Path(values[0])
        ttl_path = Path(MACROS_DIR.get()) / relative_path
        return ttl_path.resolve()
    except Exception as e:
        print(f"[パス取得エラー] {e}")
        return None

def on_double_click(event=None):
    ttl_path = get_selected_ttl_path()
    if ttl_path:
        run_ttl(ttl_path, TTERM_PATH.get())

def on_enter_key(event):
    on_double_click()

def on_right_click(event):
    ttl_path = get_selected_ttl_path()
    if ttl_path:
        try:
            subprocess.Popen(["notepad", str(ttl_path)])
        except Exception as e:
            messagebox.showerror("エラー", f"エディタ起動に失敗しました:\n{e}")

def edit_selected_ttl():
    ttl_path = get_selected_ttl_path()
    if ttl_path:
        try:
            subprocess.Popen([PREFERRED_EDITOR, str(ttl_path)])
        except Exception as e:
            messagebox.showerror("エディタ起動に失敗しました:\n{e}")

# --- ツリー構築 ---
def build_tree(tree):
    tree.delete(*tree.get_children())
    macro_root = Path(MACROS_DIR.get())
    grouped_nodes = {}
    ungrouped_ttls = []

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
                    rel_path = str(ttl_path.relative_to(macro_root))
                    tree.insert(parent, "end", iid=node_id, text=part, values=[rel_path])
                else:
                    tree.insert(parent, "end", iid=node_id, text=f"📁 {part}")
            parent = node_id

    if ungrouped_ttls:
        if not tree.exists('ungrouped'):
            tree.insert('', 'end', iid='ungrouped', text='📁 未分類')
        for ttl_path in ungrouped_ttls:
            leaf_id = f"ungrouped/{ttl_path.name}"
            rel_path = str(ttl_path.relative_to(macro_root))
            tree.insert('ungrouped', "end", iid=leaf_id, text=ttl_path.name, values=[rel_path])

# --- GUIレイアウト構築 ---
config = load_launcher_config()
TTERM_PATH.set(config.get("teraterm_path", ""))
MACROS_DIR.set(config.get("macros_root", ""))

frame_config = tk.Frame(root)
frame_config.pack(fill=tk.X, padx=10, pady=5)

# パス入力欄
labels = ["Tera Termのパス:", "TTLマクロルート:"]
entries = [TTERM_PATH, MACROS_DIR]
buttons = [
    lambda: TTERM_PATH.set(filedialog.askopenfilename(filetypes=[("実行ファイル", "*.exe")])),
    lambda: MACROS_DIR.set(filedialog.askdirectory())
]

for i in range(2):
    tk.Label(frame_config, text=labels[i]).grid(row=i, column=0, sticky="w")
    tk.Entry(frame_config, textvariable=entries[i], width=60).grid(row=i, column=1, padx=5)
    tk.Button(frame_config, text="参照", command=buttons[i]).grid(row=i, column=2, padx=5)

tk.Button(frame_config, text="保存", command=save_config).grid(row=0, column=3, padx=5)
tk.Button(frame_config, text="再読込", command=lambda: build_tree(tree)).grid(row=1, column=3, padx=5)

frame_tree = tk.Frame(root)
frame_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

# ツリー構造表示
tree = ttk.Treeview(frame_tree, columns=("path",), show="tree headings")
tree.heading("#0", text="マクロ構成")
tree.heading("path", text="TTLマクロ格納パス（相対）")
tree.column("#0", anchor="w", width=300)
tree.column("path", anchor="w", width=500)

tree.bind("<Double-1>", on_double_click)
tree.bind("<Return>", on_enter_key)
tree.bind("<Button-3>", on_right_click)

tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar = tk.Scrollbar(frame_tree, orient="vertical", command=tree.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
tree.configure(yscrollcommand=scrollbar.set)

# 実行・編集・終了ボタン
frame_bottom = tk.Frame(root)
frame_bottom.pack(pady=10)

buttons = [
    ("接続実行", lambda: on_double_click(None)),
    ("編集", edit_selected_ttl),
    ("閉じる", root.quit)
]

for i, (label, cmd) in enumerate(buttons):
    tk.Button(frame_bottom, text=label, command=cmd).grid(row=0, column=i, padx=20)

build_tree(tree)
root.mainloop()
