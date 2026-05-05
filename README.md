# Tera Term Macro Manager

Tera Term用の `.ttl` マクロファイルを Excelベースで一括管理・生成・起動できる環境を提供するツールです。  
複数拠点・複数アカウントの接続管理を効率化します。

最新リリース: [Releases](https://github.com/enikumaah1988/ttmacro-manager/releases) — Python 不要の exe 配布版（zip）はリリースページから入手できます。

---

## 特長

- **Excel ベースで一括管理**: 接続情報を `.xlsx` 1 ファイルで保守。グループ階層別に `.ttl` を自動生成（ファイル名は `論理名_IP_ユーザ名.ttl`）
- **用途別テンプレート**: Excel の `template` 列で行ごとに `macros/templates/<name>.ttl` を切り替え。空欄なら `default.ttl`
- **GUI ランチャー（ダークテーマ）**: 検索ボックス・サーバ名 / IP / ユーザ名のツリー表示・接続 / 編集ボタン
- **接続後の自動コマンド実行 + メモ機能**: 接続直後に `post_cmd` を sendln で投入、`memo` は TTL ヘッダにコメントとして残る
- **鍵認証 / パスワード認証 両対応**
- **exe 形式の配布に対応**: Python 未導入環境でも単体配布可能（[Releases](https://github.com/enikumaah1988/ttmacro-manager/releases)）
- **環境変数によるパス上書き**: `TTMACRO_*` で BASE_DIR / EXCEL_PATH 等を個別変更可能
- **`--clean` オプション**: グループ変更で残った旧 TTL を一括掃除（`--dry-run` で削除候補確認）

---

## 構成イメージ

```text
ttmacro-manager/
├── data/
│   ├── servers_template.xlsx   # 公開用テンプレート
│   └── servers.xlsx            # 実運用ファイル（Git管理外）
├── keys/                       # 鍵ファイル格納ディレクトリ
│   └── xxxx.key                # サーバ認証鍵（Git管理外）
├── macros/                     # .ttl 出力ディレクトリ（グループ別）（テンプレート以外はGit管理外）
│   ├── templates/
│   │   ├── default.ttl         # デフォルトテンプレート（template列が空の行で使用）
│   │   └── *.ttl               # 用途別テンプレート（任意）
│   └── [group1]/               # グループ別ディレクトリ
│       └── [group2]/
│           └── [group3]/
├── logs/
│   ├── generate.log            # 生成スクリプトのログ（Git管理外）
│   └── XXXXX.log               # ttl実行時のログ（Git管理外）
├── src/
│   └── ttmacro/                # 生成器・ランチャーの本体パッケージ
│       ├── config.py           # パス定数（TTMACRO_* 環境変数で上書き可）
│       ├── logger.py           # ログ設定
│       ├── path_resolver.py    # 出力先・相対パス計算
│       ├── ttl_renderer.py     # テンプレート展開
│       ├── excel_loader.py     # Excel I/O・行検証
│       ├── cleaner.py          # --clean 時の旧 TTL 削除
│       ├── cli.py              # ttmacro-generate のエントリ
│       └── launcher.py         # ttmacro-launcher のエントリ（GUI）
├── tests/                      # 単体テスト
│   ├── test_cleaner.py
│   ├── test_config.py
│   ├── test_excel_loader.py
│   ├── test_path_resolver.py
│   └── test_ttl_renderer.py
├── packaging/                  # PyInstaller 用資材
│   ├── entry_generate.py       # CLI 用エントリスクリプト
│   ├── entry_launcher.py       # GUI 用エントリスクリプト
│   ├── ttmacro-generate.spec   # CLI 用 .spec
│   ├── ttmacro-launcher.spec   # GUI 用 .spec
│   └── launcher.ico            # exe・ウィンドウ用アイコン
├── bin/
│   ├── ttmacro-generate.exe          # CLI（PyInstaller --onefile 成果物。Git管理外）
│   ├── ttmacro-launcher/             # GUI（PyInstaller --onedir 成果物のフォルダ。Git管理外）
│   │   ├── ttmacro-launcher.exe      #   実行ファイル本体
│   │   └── _internal/                #   依存ファイル一式
│   └── launcher_config.json          # ランチャーの設定ファイル（Git管理外）
├── build.ps1                   # exe ビルド + zip 化を自動化するスクリプト
├── pyproject.toml              # 依存・ツール設定
├── .gitignore
└── README.md
```

---

## セットアップ手順（Windows）

> exe 配布版を使うだけの方は、[Releases](https://github.com/enikumaah1988/ttmacro-manager/releases) から zip を取得して解凍するだけで動きます。以下はソースから動かす開発者向け手順です。

### 1. Pythonをインストール（初回のみ）

[公式サイト](https://www.python.org/downloads/windows/)から最新版をインストール  
※ インストール時に「**Add Python to PATH**」にチェックを忘れずに！

### 2. リポジトリを取得（初回のみ）

```powershell
git clone https://github.com/enikumaah1988/ttmacro-manager.git
cd ttmacro-manager
```

### 3. 仮想環境（.venv）の利用

**初回のみ** — 仮想環境を作成してから有効化：

```powershell
# Python 3.12 以上で作成（複数バージョンがある場合は py -3.12 -m venv .venv のように明示指定）
python -m venv .venv
.venv\Scripts\Activate
```

**2回目以降** — 仮想環境を有効にするだけ：

```powershell
.venv\Scripts\Activate
```

仮想環境が有効になると、プロンプトの先頭に `(.venv)` が付きます。  
表示されない場合は、Windowsの設定（実行ポリシー等）を確認してください。

```text
(.venv) C:\path\to\ttmacro-manager>
```

---

### 4. プロジェクトをインストール

プロジェクトを編集可能な形で `pip install` してください：

```powershell
# 通常はこちら（開発ツール込み）
pip install -e ".[dev]"

# 動作だけ確認したい場合
pip install -e .
```

インストール後、仮想環境（`.venv`）を Activate した状態であれば、`ttmacro-generate` と `ttmacro-launcher` の 2 つをコマンドとして直接実行できます。（実体は `.venv\Scripts\` 配下に作られる shim exe です。Activate すると同フォルダがシェルの検索パス先頭に入るためコマンド名だけで起動できます。）

Activate していない場合は `.venv\Scripts\ttmacro-generate.exe` のように直接フルパスで呼ぶこともできます。

---

### 5. Excel台帳ファイルの準備

以下のようなExcelファイルを `data/servers.xlsx` として用意します。  
`servers_template.xlsx` をコピーして編集してください：

```powershell
copy data\servers_template.xlsx data\servers.xlsx
```

#### Excelファイルの構成（`servers.xlsx`）

| group1    | group2 | group3 | name    | host         | port | user  | password | keyfile        | post_cmd | generate | template | memo       |
|-----------|--------|--------|---------|--------------|------|-------|----------|----------------|----------|----------|----------|------------|
| LocationA | NAS    | 管理   | infra01 | 192.168.0.10 | 22   | rocky | rocky123 |                | date     | yes      |          | 本社NAS    |
| LocationB |        |        | infra02 | 192.168.0.11 | 22   | rocky |          | id_ed25519.ppk |          | yes      |          | 支社サーバ |
| LocationC | 開発   | テスト | dev01   | 192.168.0.12 | 22   | dev   | devpass  |                | ls -la   | yes      | nodejs   | 開発サーバ |
| LocationC | 開発   | テスト | dev02   | 192.168.0.13 | 22   | dev   | devpass  |                | ls -la   | yes      | nodejs   | 開発サーバ |
|           |        |        |         |              |      |       |          |                |          |          |          |            |

- `generate` 列が `yes` の行だけが `.ttl` 生成対象になります。
- `e` を検出した時点で処理終了
- 空白行はスキップされます
- `keyfile` は `keys/` ディレクトリ内のファイル名を記載してください。
- `group1` が空欄の場合は `macros/` に直下出力されます。子グループだけの指定は無効です。
- `post_cmd` は接続後に自動実行するコマンドを記載します。複数行の場合は改行で区切ります。
- `template` は使用するテンプレート名（拡張子省略可）。空欄なら `default.ttl`、`nodejs` と書けば `macros/templates/nodejs.ttl` を使う。
- `memo` は接続情報のメモを記載します。TTLファイルのヘッダーに表示されます。

---

### 6. TTLマクロを生成

```powershell
# 全行を生成（generate=yes の行のみ対象）
ttmacro-generate

# 特定の No. のみ生成（generate 列の値に関係なく対象行を処理）
ttmacro-generate --row 5

# 既存 TTL を全削除してから生成（グループ変更で残った旧 TTL の掃除）
ttmacro-generate --clean

# 削除対象を確認するだけ（実削除も生成もしない）
ttmacro-generate --clean --dry-run
```

`template` 列に値を入れた行は `macros/templates/<その値>.ttl` を、空欄の行は `default.ttl` を使って生成されます。

---

### 7. TTLを選んで起動

```powershell
ttmacro-launcher
```

ランチャーの初回起動後、画面上の「参照」ボタンで Tera Term の実行ファイルパスと macros ルートを指定し「設定保存」を押すと `bin/launcher_config.json` に自動保存されます（手書きで作成する必要はありません）。

---

## 🖼 GUIランチャー画面イメージ

![Tera Term GUIランチャー](images/launcher_gui.png)

ダークテーマで、検索ボックス・サーバ名 / IP / ユーザ名の 3 列表示・接続 / 編集 / 閉じるボタンを備える。

---

## ログファイルのフォーマット

### `logs/generate.log` — TTL 生成時のログ

```text
2024-03-15 14:30:22 - 生成開始
2024-03-15 14:30:22 - ✅ infra01_192.168.0.10_rocky.ttl を生成しました。
2024-03-15 14:30:22 - ✅ infra02_192.168.0.11_rocky.ttl を生成しました。
2024-03-15 14:30:22 - ✅ dev01_192.168.0.12_dev.ttl を生成しました。
2024-03-15 14:30:22 - ⏹️ 'e' を検出したため、処理を終了します。
```

### `logs/{論理名}_{IP}_{ユーザ名}_{YYYYMMDD_HHMMSS}.log` — TTL 実行時のログ

例: `logs/infra01_192.168.0.10_rocky_20240315_143022.log`

```text
[2024-03-15 14:30:22.123] [rocky@infra01 ~]$ date '+%Y/%m/%d %H:%M:%S'
2024/03/15 14:30:23
[2024-03-15 14:30:23.456] [rocky@infra01 ~]$ whoami
rocky
[2024-03-15 14:30:23.789] [rocky@infra01 ~]$ uname -a
Linux infra01 5.15.0-91-generic #101-Ubuntu SMP Tue Nov 14 13:30:08 UTC 2023 x86_64 x86_64 x86_64 GNU/Linux
```

---

## TTL ファイル運用上の注意

### 生成済み TTL を別の階層に移動しないでください

TTL ファイルは生成時に **配置場所からプロジェクトルートまでの相対パス** を内部に埋め込んで保存します。これによりログ・鍵ファイルが正しく解決されます。生成後に `.ttl` を別フォルダへ手動で移動するとパス解決が壊れます。グループ階層を変えたいときは、Excel の `group1` 〜 `group3` 列を変更して `ttmacro-generate --clean` で再生成してください。

### Tera Term のログ設定について

Tera Term の `teraterm.ini` にログ設定が記述されていると、TTL マクロ内の `logopen` 指定よりも **INI 側の設定が優先** されます。プロジェクト側で指定した `logs/` 配下にログを残したい場合は、`teraterm.ini` のログ関連項目を空欄にしておく必要があります。

---

## トラブルシューティング

### よくある問題と解決方法

1. **TTLファイルが見つからない**
   - 確認事項：
     - `macros_root`の設定が正しいか
     - TTLファイルが正しいディレクトリに生成されているか
   - 解決方法：
     - `launcher_config.json`の`macros_root`を確認
     - `ttmacro-generate` を再実行

2. **鍵ファイルが見つからない**
   - 確認事項：
     - `keys/`ディレクトリに鍵ファイルが存在するか
     - Excelファイルの`keyfile`列の値が正しいか
   - 解決方法：
     - 鍵ファイルを`keys/`ディレクトリに配置
     - Excelファイルの`keyfile`列を確認

3. **ログファイルが生成されない**
   - 確認事項：
     - `logs/` ディレクトリが存在するか
     - 書き込み権限があるか
   - 解決方法：
     - `logs/` ディレクトリを手動で作成するか、生成スクリプト実行時に自動作成されることを確認
     - ディレクトリの権限を確認

4. **パス解決エラー**
   - 確認事項：
     - TTLファイルの配置場所が正しいか
     - グループ階層が正しく設定されているか
   - 解決方法：
     - TTLファイルを正しいディレクトリに移動
     - Excelファイルのグループ設定を確認

5. **Pythonの実行エラー / コマンドが見つからない**
   - 確認事項：
     - 仮想環境（.venv）が Activate 済みか（プロンプト先頭に `(.venv)` が付いているか）
     - 必要なライブラリがインストールされているか
     - `ttmacro-generate` / `ttmacro-launcher` の shim exe が `.venv\Scripts\` 配下に存在するか
   - 解決方法：
     - `.venv\Scripts\activate` を実行（これでシェルの検索パスに `Scripts\` が入りコマンドが効く）
     - `pip install -e ".[dev]"` を実行（インストール時に shim exe が `.venv\Scripts\` に作られる）
     - Activate しない場合は `.venv\Scripts\ttmacro-generate.exe` のように直接フルパスで呼ぶ

6. **Excelファイルの読み込みエラー**
   - 確認事項：
     - Excelファイルが他のアプリケーションで開かれていないか
     - ファイルの形式が正しいか（.xlsx）
   - 解決方法：
     - Excelファイルを閉じる
     - ファイルを.xlsx形式で保存し直す

## セキュリティに関する注意事項

- **パスワードを Excel に平文で書かない**: 可能な限り鍵認証（`keyfile` 列）を使う。やむを得ず `password` 列を使う場合は `servers.xlsx` のアクセス権を絞る
- **生成物は接続情報を含む**: `keys/` `macros/**/*.ttl` `logs/*.log` には接続先・パスワード・実行履歴が残るため、git にはコミットせず（既に `.gitignore` 済み）、共有ストレージや配布物に含めないよう注意

---

## exe 配布版のビルド（Python 不要環境向け）

PyInstaller で `bin/*.exe` を生成して、Python が入っていない端末にも配布できます。

### おすすめ: build.ps1 でまとめて作成

プロジェクトルートで以下を実行するだけで、exe ビルド → 配布フォルダ組み立て → zip 化まで自動化されます。

```powershell
.\build.ps1
```

成果物:

- `bin/ttmacro-launcher/`（GUI、フォルダ約 45 MB。中の `ttmacro-launcher.exe` を起動）
- `bin/ttmacro-generate.exe`（CLI、約 16 MB）
- `deploy/`（配布用フォルダ。`bin/` `data/` `macros/templates/` `keys/` `logs/` を含む）
- `ttmacro-manager-v<version>.zip`（`deploy/` を圧縮したもの。配布相手に渡す）

GUI ランチャーは `--onedir` モードでビルドしているため、単一 exe ではなくフォルダ単位の配布になります（起動時のファイル展開を省くため、起動が `--onefile` の 3 倍以上速い）。

配布相手は zip を解凍 → `data/servers_template.xlsx` を `servers.xlsx` にリネーム + 編集 → `bin/ttmacro-launcher/ttmacro-launcher.exe` を起動、で使い始められます。

### 個別にビルドしたい場合

```powershell
# GUI ランチャー
.venv\Scripts\pyinstaller --clean --distpath bin packaging/ttmacro-launcher.spec

# TTL 生成 CLI
.venv\Scripts\pyinstaller --clean --distpath bin packaging/ttmacro-generate.spec
```

exe は `<deploy_root>/bin/` 配下に置く前提で、`data/` `macros/` 等を相対参照します。`bin/` 単体で別フォルダに移動すると動かないので注意してください。
