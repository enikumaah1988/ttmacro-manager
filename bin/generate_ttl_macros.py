import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import math
import logging
import argparse
from typing import Dict, List, Optional

# 各種パスの定義
BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_PATH = BASE_DIR / "data" / "servers.xlsx"
TEMPLATE_PATH = BASE_DIR / "macros" / "template.ttl"
OUTPUT_DIR = BASE_DIR / "macros"
LOGS_DIR = BASE_DIR / "logs"
KEYS_DIR = BASE_DIR / "keys"

# ログ設定
def setup_logging():
    """ログ設定を行う"""
    log_file = LOGS_DIR / "generate.log"
    LOGS_DIR.mkdir(exist_ok=True)
    
    # ログフォーマットの設定
    formatter = logging.Formatter('[%(asctime)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    
    # ファイルハンドラの設定
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setFormatter(formatter)
    
    # 標準出力ハンドラの設定
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    
    # ロガーの設定
    logger = logging.getLogger('generate')
    logger.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

# TTLテンプレート読み込み
def load_template() -> str:
    """TTLテンプレートを読み込む"""
    return TEMPLATE_PATH.read_text(encoding="utf-8")

def sanitize_name(name: str) -> str:
    """Windows禁止文字を _ に置換"""
    return re.sub(r'[\\/:*?"<>|]', '_', name)

def safe_str(val) -> str:
    """ExcelのNaNを空文字に変換"""
    if isinstance(val, float) and math.isnan(val):
        return ""
    return str(val).strip()

def safe_get(row: pd.Series, key: str, default: str = "") -> str:
    """行データから安全に値を取得し、NaNの場合はデフォルト値を返す"""
    value = row.get(key, default)
    return str(value if pd.notna(value) else default).strip()

def load_excel_data() -> pd.DataFrame:
    """Excelファイルを読み込む"""
    try:
        with open(EXCEL_PATH, 'rb') as f:
            return pd.read_excel(f, engine="openpyxl")
    except PermissionError:
        print(f"⚠️ Excelファイルが他で開かれています: {EXCEL_PATH}")
        print("💡 閉じてから再度実行してください。")
        exit(1)

def extract_row_data(row: pd.Series) -> Dict[str, str]:
    """行データから必要な情報を抽出"""
    # 特殊処理が必要なフィールド
    memo = safe_get(row, "memo").replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
    
    return {
        "name": sanitize_name(row["name"]),
        "host": row["host"],
        "port": str(row["port"]),
        "user": row["user"],
        "password": safe_get(row, "password"),
        "keyfile_name": safe_get(row, "keyfile"),
        "post_cmd": safe_get(row, "post_cmd"),
        "memo": memo,
        "group1": safe_get(row, "group1"),
        "group2": safe_get(row, "group2"),
        "group3": safe_get(row, "group3")
    }

def get_target_directory(data: Dict[str, str]) -> Path:
    """グループ階層に基づいて出力ディレクトリを決定"""
    if not data["group1"]:
        return OUTPUT_DIR
    
    target_dir = OUTPUT_DIR / data["group1"]
    if data["group2"]:
        target_dir = target_dir / data["group2"]
        if data["group3"]:
            target_dir = target_dir / data["group3"]
    
    target_dir.mkdir(parents=True, exist_ok=True)
    return target_dir

def calculate_relative_path(target_dir: Path) -> str:
    """TTLファイルの配置場所からプロジェクトルートへの相対パスを計算"""
    # プロジェクトルートからの相対パスを計算
    rel_path = target_dir.relative_to(BASE_DIR)
    
    # 相対パスを文字列に変換し、必要に応じて'../'を追加
    if rel_path == Path('.'):
        return ''
    
    # ディレクトリの深さに応じて'../'を追加
    depth = len(rel_path.parts)
    return '../' * depth

def calculate_paths(data: Dict[str, str], target_dir: Path) -> Dict[str, str]:
    """各種パスを計算"""
    # TTLファイル名の生成
    ttl_name = f"{data['name']}_{data['user']}_{data['host']}"
    
    # ログファイル名の生成
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f"{ttl_name}_{timestamp}.log"
    log_file = LOGS_DIR / log_filename
    
    # キーファイルパスの計算
    keyfile_path = ""
    if data["keyfile_name"]:
        keyfile_path = str(KEYS_DIR / data["keyfile_name"])
    
    return {
        "ttl_name": ttl_name,
        "log_file": str(log_file),
        "log_path": str(LOGS_DIR),
        "keyfile": keyfile_path
    }

def generate_ttl_content(data: Dict[str, str], template: str, timestamp: str, target_dir: Path) -> str:
    """TTLマクロの内容を生成"""
    # 相対パスの計算
    rel_path = calculate_relative_path(target_dir)
    
    # ポストコマンドの処理
    post_cmd_lines = [line.strip() for line in data["post_cmd"].splitlines() if line.strip()]
    post_commands = "\n".join([
        f"wait '$' '#'\nsendln '{cmd}'\n" for cmd in post_cmd_lines
    ]) if post_cmd_lines else ""
    
    # テンプレートの置換
    replacements = {
        "{hostname}": data["host"],
        "{port}": data["port"],
        "{username}": data["user"],
        "{password}": data["password"],
        "{keyfile}": data["keyfile_name"],  # キーファイル名のみを渡す
        "{name}": data["name"],
        "{rel_path}": rel_path,  # 相対パスを渡す
        "{created_at}": timestamp,
        "{memo}": data["memo"],
        "{post_commands}": post_commands
    }
    
    content = template
    for key, value in replacements.items():
        content = content.replace(key, value)
    
    return content

def parse_args():
    """コマンドライン引数を解析"""
    parser = argparse.ArgumentParser(
        description=r'''
TTLマクロを生成するツール

servers.xlsxの内容に基づいてTTLマクロを生成します。
generate列が'yes'の行のみが処理対象となります。
'e'が指定されている行で処理を終了します。

実行方法:
  # PowerShellの場合
  python .\generate_ttl_macros.py [オプション]

  # コマンドプロンプトの場合
  python generate_ttl_macros.py [オプション]
        ''',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=r'''
使用例:
  # 全行を生成
  python .\generate_ttl_macros.py

  # 特定の行のみ生成（5行目）
  python .\generate_ttl_macros.py --row 5

  # ヘルプを表示
  python .\generate_ttl_macros.py --help

注意:
  - 行番号はExcelのA列のNo.を指定します
  - generate列が'yes'の行のみが処理されます
  - 生成フラグに'e'を指定すると処理を終了します
  - PowerShellで実行する場合は 'python .\generate_ttl_macros.py' を使用してください
        '''
    )
    parser.add_argument(
        '--row', 
        type=int, 
        help='生成する行番号（1から始まる）。指定がない場合は全行を処理します。'
    )
    return parser.parse_args()

def generate_ttl_macros(args):
    """TTLマクロを生成するメイン関数"""
    logger = setup_logging()
    template = load_template()
    df = load_excel_data()
    timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    
    logger.info("生成開始")
    
    # 行番号が指定されている場合
    if args.row is not None:
        # No.列で指定された行を検索
        matching_rows = df[df['No.'] == args.row]
        if matching_rows.empty:
            logger.error(f"❌ 指定されたNo. {args.row} は見つかりませんでした")
            return
        # 指定された行のみ処理
        rows_to_process = [matching_rows.iloc[0]]
        logger.info(f"📝 No.{args.row} のサーバーを処理します")
    else:
        # 全行処理
        rows_to_process = df.iterrows()
    
    for row in rows_to_process:
        # iterrowsの場合はタプルが返るので、行データを取得
        if isinstance(row, tuple):
            _, row = row
        
        # 空白行スキップ
        if row.isnull().all():
            continue
        
        # 生成フラグを確認
        generate_flag = str(row.get("generate", "")).strip().lower()
        if generate_flag == "e":
            logger.info("⏹️ 'e' を検出したため、処理を終了します。")
            break
        if generate_flag != "yes":
            continue
        
        try:
            # データの抽出と処理
            data = extract_row_data(row)
            target_dir = get_target_directory(data)
            content = generate_ttl_content(data, template, timestamp, target_dir)
            
            # マクロファイルを生成
            ttl_name = f"{data['name']}_{data['user']}_{data['host']}"
            (target_dir / f"{ttl_name}.ttl").write_text(content, encoding="utf-8")
            logger.info(f"✅ {ttl_name}.ttl を生成しました。（No.{row['No.']}）")
        except Exception as e:
            logger.error(f"❌ {ttl_name}.ttl の生成に失敗しました: {str(e)}")

if __name__ == "__main__":
    args = parse_args()
    generate_ttl_macros(args)
    exit(0)