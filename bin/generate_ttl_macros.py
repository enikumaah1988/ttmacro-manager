# 型ヒントを遅延評価にし、pandas を後からインポートできるようにする
from __future__ import annotations

# 最初に使用中の Python を表示（pandas は後でインポートするのでここでは落ちない）
import sys
print("使用中の Python:", sys.executable, file=sys.stderr, flush=True)
# Python 3.14 では pandas/numpy のネイティブ拡張が未対応で import 時に落ちるためチェック
print("TTLマクロ生成を開始しています...", file=sys.stderr, flush=True)

# pandas は generate_ttl_macros() 内で遅延インポート（import で落ちる環境でもスクリプトは起動する）
pd = None

from pathlib import Path
from datetime import datetime
import re
import math
import logging
import argparse
import traceback
import ipaddress
from typing import Dict, List, Optional, Tuple

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
    
    # コンソールハンドラ（stderr に明示）
    console_handler = logging.StreamHandler(sys.stderr)
    console_handler.setFormatter(formatter)
    
    # ロガーの設定（既存ハンドラをクリアしてから追加）
    logger = logging.getLogger('generate')
    logger.handlers.clear()
    logger.setLevel(logging.INFO)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    logger.propagate = False
    
    return logger

# TTLテンプレート読み込み
def load_template() -> str:
    """TTLテンプレートを読み込む"""
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"テンプレートファイルが見つかりません: {TEMPLATE_PATH}")
    
    try:
        content = TEMPLATE_PATH.read_text(encoding="utf-8")
        if not content.strip():
            raise ValueError("テンプレートファイルが空です")
        return content
    except UnicodeDecodeError:
        raise ValueError(f"テンプレートファイルの文字エンコーディングが不正です: {TEMPLATE_PATH}")
    except Exception as e:
        raise RuntimeError(f"テンプレートファイル読み込みエラー: {str(e)}")

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
    if not EXCEL_PATH.exists():
        raise FileNotFoundError(f"Excelファイルが見つかりません: {EXCEL_PATH}")
    
    try:
        with open(EXCEL_PATH, 'rb') as f:
            df = pd.read_excel(f, engine="openpyxl")
            if df.empty:
                raise ValueError("Excelファイルが空です")
            return df
    except PermissionError:
        raise PermissionError(f"Excelファイルが他で開かれています: {EXCEL_PATH}")
    except Exception as e:
        raise RuntimeError(f"Excelファイル読み込みエラー: {str(e)}")

def validate_row_data(row: pd.Series, row_num: int) -> Tuple[bool, List[str]]:
    """行データの妥当性を検証"""
    errors = []
    
    # 必須フィールドチェック
    required_fields = ['name', 'host', 'user']
    for field in required_fields:
        if pd.isna(row.get(field)) or str(row.get(field, '')).strip() == '':
            errors.append(f"必須項目 '{field}' が空です")
    
    # IPアドレス/ホスト名チェック
    host = str(row.get('host', '')).strip()
    if host:
        try:
            ipaddress.ip_address(host)
        except ValueError:
            # IPアドレスでない場合はホスト名として扱う（簡易チェック）
            if not re.match(r'^[a-zA-Z0-9.-]+$', host):
                errors.append(f"ホスト名 '{host}' の形式が不正です")
    
    # ポート番号チェック
    port = row.get('port')
    if pd.notna(port):
        try:
            port_num = int(port)
            if not (1 <= port_num <= 65535):
                errors.append(f"ポート番号 {port_num} は範囲外です (1-65535)")
        except (ValueError, TypeError):
            errors.append(f"ポート番号 '{port}' が数値ではありません")
    
    # キーファイル存在チェック
    keyfile = safe_get(row, 'keyfile')
    if keyfile:
        keyfile_path = KEYS_DIR / keyfile
        if not keyfile_path.exists():
            errors.append(f"キーファイル '{keyfile}' が見つかりません: {keyfile_path}")
    
    return len(errors) == 0, errors

def extract_row_data(row: pd.Series) -> Dict[str, str]:
    """行データから必要な情報を抽出"""
    # 特殊処理が必要なフィールド
    memo = safe_get(row, "memo").replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
    
    return {
        "name": sanitize_name(str(row["name"]).strip()),
        "host": str(row["host"]).strip(),
        "port": str(int(row["port"])) if pd.notna(row["port"]) else "22",
        "user": str(row["user"]).strip(),
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
    
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
        # 書き込み権限チェック
        test_file = target_dir / ".write_test"
        test_file.touch()
        test_file.unlink()
    except PermissionError:
        raise PermissionError(f"ディレクトリへの書き込み権限がありません: {target_dir}")
    except Exception as e:
        raise RuntimeError(f"ディレクトリ作成エラー: {target_dir} - {str(e)}")
    
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

def get_log_dir(target_dir: Path) -> Path:
    """TTL と同じ階層になるよう logs 以下のディレクトリを返す（macros/home/prod → logs/home/prod）"""
    rel = target_dir.relative_to(OUTPUT_DIR)
    if rel == Path("."):
        return LOGS_DIR
    return LOGS_DIR / rel


def calculate_paths(data: Dict[str, str], target_dir: Path) -> Dict[str, str]:
    """各種パスを計算"""
    # TTLファイル名の生成
    ttl_name = f"{data['name']}_{data['host']}_{data['user']}"
    
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
    global pd
    # ここで pandas をインポート（import で落ちる環境でもスクリプトはここまで起動する）
    if pd is None:
        try:
            import pandas as _pd
            pd = _pd
        except Exception as e:
            print("pandas のインポートに失敗しました:", e, file=sys.stderr, flush=True)
            print("仮想環境を有効にして、pip install pandas openpyxl を実行してください。", file=sys.stderr, flush=True)
            sys.exit(1)
    print("[1/4] ログ設定...", file=sys.stderr, flush=True)
    logger = setup_logging()
    
    try:
        # 初期化処理
        print("[2/4] テンプレート・Excel 読み込み...", file=sys.stderr, flush=True)
        template = load_template()
        df = load_excel_data()
        timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        
        logger.info(f"読み込み元: {EXCEL_PATH}")
        logger.info("生成開始")
        
        # 必要な列の存在チェック
        required_columns = ['No.', 'name', 'host', 'user', 'generate']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"必要な列が見つかりません: {', '.join(missing_columns)}")
        
        # 行番号が指定されている場合
        if args.row is not None:
            matching_rows = df[df['No.'] == args.row]
            if matching_rows.empty:
                logger.error(f"❌ 指定されたNo. {args.row} は見つかりませんでした")
                return
            rows_to_process = [(args.row, matching_rows.iloc[0])]
            logger.info(f"📝 No.{args.row} のサーバーを処理します")
        else:
            rows_to_process = df.iterrows()
            # 対象行数を事前に表示（generate=yes の行数）
            generate_count = sum(
                1 for _, r in df.iterrows()
                if str(r.get("generate", "")).strip().lower() in ("yes", "true", "1")
            )
            logger.info(f"generate=yes の行: {generate_count} 件（全 {len(df)} 行中）")
            if generate_count == 0:
                logger.warning("⚠️ 対象行が0件です。Excelの generate 列に yes を指定した行がありますか？")
        
        print("[3/4] 行を処理しています...", file=sys.stderr, flush=True)
        success_count = 0
        error_count = 0
        
        for idx, row in rows_to_process:
            try:
                # 空白行スキップ
                if row.isnull().all():
                    continue
                
                # 生成フラグを確認（--row 指定時は対象行を無条件で処理）
                generate_flag = str(row.get("generate", "")).strip().lower()
                if args.row is None and generate_flag == "e":
                    logger.info("⏹️ 'e' を検出したため、処理を終了します。")
                    break
                # yes/true/1 を有効とする（Excel の TRUE や 1 にも対応）
                if args.row is None and generate_flag not in ("yes", "true", "1"):
                    continue
                
                # 行データの検証
                row_num = row.get('No.', idx + 1)
                is_valid, validation_errors = validate_row_data(row, row_num)
                if not is_valid:
                    error_msg = f"No.{row_num} データ検証エラー: {'; '.join(validation_errors)}"
                    logger.error(f"❌ {error_msg}")
                    error_count += 1
                    continue
                
                # データの抽出と処理
                data = extract_row_data(row)
                target_dir = get_target_directory(data)
                content = generate_ttl_content(data, template, timestamp, target_dir)
                
                # マクロファイルを生成
                ttl_name = f"{data['name']}_{data['host']}_{data['user']}"
                ttl_file = target_dir / f"{ttl_name}.ttl"
                
                try:
                    ttl_file.write_text(content, encoding="utf-8")
                    logger.info(f"✅ {ttl_name}.ttl を生成しました。（No.{row_num}）")
                    success_count += 1
                except Exception as e:
                    logger.error(f"❌ ファイル書き込みエラー {ttl_name}.ttl: {str(e)}")
                    error_count += 1
                    
            except Exception as e:
                row_num = row.get('No.', idx + 1) if not row.isnull().all() else idx + 1
                logger.error(f"❌ No.{row_num} 処理エラー: {str(e)}")
                error_count += 1
        
        # 処理結果サマリー
        print("[4/4] 完了", file=sys.stderr, flush=True)
        logger.info(f"📊 処理完了 - 成功: {success_count}件, エラー: {error_count}件")
        
    except Exception as e:
        err_msg = f"致命的エラー: {str(e)}"
        tb_lines = traceback.format_exc()
        # 必ず stderr にトレースバックを出す（ログ未初期化でも確実に表示）
        print("", file=sys.stderr)
        print("=== エラー内容（トレースバック） ===", file=sys.stderr)
        print(tb_lines, file=sys.stderr)
        print("====================================", file=sys.stderr)
        try:
            logger.error(f"❌ {err_msg}")
            # ログファイルにもトレースバックを残す（コンソールに出ない場合のため）
            logger.error("トレースバック:\n%s", tb_lines)
        except NameError:
            print(f"エラー: {err_msg}", file=sys.stderr)
        sys.stderr.flush()
        sys.exit(1)


if __name__ == "__main__":
    try:
        args = parse_args()
        generate_ttl_macros(args)
        print("TTLマクロ生成を終了しました。", file=sys.stderr, flush=True)
        sys.exit(0)
    except SystemExit:
        raise
    except Exception:
        # どこで落ちてもトレースバックを必ず stderr に出す
        tb_lines = traceback.format_exc()
        print("", file=sys.stderr)
        print("=== 予期しないエラー（トレースバック） ===", file=sys.stderr)
        print(tb_lines, file=sys.stderr)
        print("==========================================", file=sys.stderr)
        sys.stderr.flush()
        # コンソールに出ない場合に備え、クラッシュログをファイルに残す
        try:
            crash_log = BASE_DIR / "logs" / "generate_crash.log"
            crash_log.parent.mkdir(parents=True, exist_ok=True)
            crash_log.write_text(tb_lines, encoding="utf-8")
        except Exception:
            pass
        sys.exit(1)