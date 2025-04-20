# Tera Term Macro Manager

Tera Term用の `.ttl` マクロファイルを Excelベースで一括管理・生成・起動できる環境を提供するツールです。  
複数拠点・複数アカウントの接続管理を効率化します。

---

## 特長

- 接続情報を **Excelファイル（.xlsx）で管理**
- `generate=yes` の行だけ `.ttl` マクロを自動生成
- グループ別ディレクトリに `.ttl` を出力（`macros/home/` など）
- ユーザー名・ホスト名・論理名を含めた明確なファイル名で生成
- ログ（`logs/generate.log`）で生成・エラーを一元記録
- Pythonスクリプトで **WindowsのTera Termと連携可能**

---

## 構成イメージ

```
ttmacro-manager/
├── data/
│   ├── servers_template.xlsx   # 公開用テンプレート
│   └── servers.xlsx            # 実運用ファイル（Git管理外）
├── keys/                       # 鍵ファイル（Git管理外）
├── macros/                     # .ttl 出力ディレクトリ（グループ別）
├── logs/
│   ├── generate.log            # ログファイル（生成・エラーを記録）
│   └── XXXXX.log               # ttl実行時のログ
├── bin/
│   ├── generate_ttl_macros.py  # TTLマクロ生成スクリプト
│   ├── run_launcher.py         # TTLを選んで接続
│   └── config.ini              # run_launcher.pyの設定ファイル
├── requirements.txt            # ライブラリ一覧
├── .gitignore
└── README.md
```

---

## セットアップ手順（Windows）

### 1. Pythonをインストール（初回のみ）

[公式サイト](https://www.python.org/downloads/windows/)から最新版をインストール  
※ インストール時に「**Add Python to PATH**」にチェックを忘れずに！

### 2. 仮想環境（.venv）を作成

```powershell
python -m venv .venv
.venv\Scripts\activate
```

仮想環境が有効になると、プロンプトが以下のようになります：

```
(.venv) C:\path\to\ttmacro-manager>
```

---

### 3. 必要なライブラリをインストール

```powershell
pip install -r requirements.txt
```

---

### 4. Excel台帳ファイルの準備

以下のようなExcelファイルを `data/servers.xlsx` として用意します。  
`servers_template.xlsx` をコピーして編集してください：

```powershell
copy data\servers_template.xlsx data\servers.xlsx
```

#### Excelファイルの構成（`servers.xlsx`）

| group1 | group2 | group3 | name     | host          | port | user   | password | keyfile          | generate |
|--------|--------|--------|----------|---------------|------|--------|----------|------------------|----------|
| 自宅   | NAS    | 管理   | infra01  | 192.168.0.10  | 22   | rocky  | rocky123 |                  | yes      |
| 実家   |        |        | infra02  | 192.168.0.11  | 22   | rocky  |          | id_ed25519.ppk   | yes      |
|        |        |        |          |               |      |        |          |                  |          |

- `generate` 列が `yes` の行だけが `.ttl` 生成対象になります。
- `e` を検出した時点で処理終了
- 空白行はスキップされます
- `keyfile` は `keys/` ディレクトリ内のファイル名を記載してください。
- `group1` が空欄の場合は `macros/` に直下出力されます。子グループだけの指定は無効です。

---

---

### 5. TTLマクロを生成

```powershell
python bin/generate_ttl_macros.py
```

---

### 6. TTLを選んで起動

```powershell
python bin/run_launcher.py
```
## 🖼 GUIランチャー画面イメージ

![Tera Term GUIランチャー](images/launcher_gui.png)

---

## 今後の展望

- コメントや用途別テンプレートの自動適用
- グループ変更時の旧ファイル削除機能
- ttlマクロ自動ログイン時のサーバ固有のカスタムコマンド発行
- パスワード暗号化
- 複数template.ttlの活用
- マクロのバージョン管理
　- 変更履歴の追跡
　- 以前のバージョンに戻す。
- UI/UXの改善

