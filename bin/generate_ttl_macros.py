import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import math
import logging
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
    return {
        "name": sanitize_name(row["name"]),
        "host": row["host"],
        "port": str(row["port"]),
        "user": row["user"],
        "password": safe_str(row.get("password", "")),
        "keyfile_name": safe_str(row.get("keyfile", "")),
        "post_cmd": safe_str(row.get("post_cmd", "")),
        "memo": str(row.get("memo", "")).strip().replace('\r', ' ').replace('\n', ' ').replace('\t', ' '),
        "group1": str(row.get("group1", "") if pd.notna(row.get("group1")) else "").strip(),
        "group2": str(row.get("group2", "") if pd.notna(row.get("group2")) else "").strip(),
        "group3": str(row.get("group3", "") if pd.notna(row.get("group3")) else "").strip()
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

def generate_ttl_content(data: Dict[str, str], template: str, timestamp: str) -> str:
    """TTLマクロの内容を生成"""
    # キーファイルのパスを生成
    keyfile = (KEYS_DIR / data["keyfile_name"]).as_posix() if data["keyfile_name"] else ""
    
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
        "{keyfile}": keyfile,
        "{name}": data["name"],
        "{ttl_name}": f"{data['name']}_{data['user']}_{data['host']}",
        "{logspath}": LOGS_DIR.resolve().as_posix() + "/",
        "{created_at}": timestamp,
        "{memo}": data["memo"],
        "{post_commands}": post_commands
    }
    
    content = template
    for key, value in replacements.items():
        content = content.replace(key, value)
    
    return content

def generate_ttl_macros():
    """TTLマクロを生成するメイン関数"""
    logger = setup_logging()
    template = load_template()
    df = load_excel_data()
    timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
    
    logger.info("生成開始")
    
    for _, row in df.iterrows():
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
            content = generate_ttl_content(data, template, timestamp)
            
            # マクロファイルを生成
            ttl_name = f"{data['name']}_{data['user']}_{data['host']}"
            (target_dir / f"{ttl_name}.ttl").write_text(content, encoding="utf-8")
            logger.info(f"✅ {ttl_name}.ttl を生成しました。")
        except Exception as e:
            logger.error(f"❌ {ttl_name}.ttl の生成に失敗しました: {str(e)}")

if __name__ == "__main__":
    generate_ttl_macros()
    exit(0)