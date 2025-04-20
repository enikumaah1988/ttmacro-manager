import pandas as pd
from pathlib import Path
from datetime import datetime

# 各種パスの定義
BASE_DIR = Path(__file__).resolve().parent.parent
EXCEL_PATH = BASE_DIR / "data" / "servers.xlsx"
TEMPLATE_PATH = BASE_DIR / "macros" / "template.ttl"
OUTPUT_DIR = BASE_DIR / "macros"
LOGS_DIR = BASE_DIR / "logs"
KEYS_DIR = BASE_DIR / "keys"

# TTLテンプレート読み込み
template = TEMPLATE_PATH.read_text(encoding="utf-8")

# Excel読み込み
try:
    with open(EXCEL_PATH, 'rb') as f:
        df = pd.read_excel(f, engine="openpyxl")
except PermissionError:
    print(f"⚠️ Excelファイルが他で開かれています: {EXCEL_PATH}")
    print("💡 閉じてから再度実行してください。")
    exit(1)

# マクロ生成
timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")
for _, row in df.iterrows():
    if row.isnull().all():
        continue  # 空白行スキップ

    generate_flag = str(row.get("generate", "")).strip().lower()
    if generate_flag == "e":
        print("⏹️ 'e' を検出したため、処理を終了します。")
        break
    if generate_flag != "yes":
        continue  # yes 以外はスキップ
    name = row["name"]
    host = row["host"]
    port = str(row["port"])
    user = row["user"]
    password = row.get("password", "") or ""
    keyfile_name = str(row.get("keyfile", "") or "").strip()
    keyfile = (KEYS_DIR / keyfile_name).as_posix() if keyfile_name else ""

    ttl_name = f"{name}_{user}_{host}"
    logspath = LOGS_DIR.resolve().as_posix() + "/"

    content = template.replace("{hostname}", host)
    content = content.replace("{port}", port)
    content = content.replace("{username}", user)
    content = content.replace("{password}", password)
    content = content.replace("{keyfile}", keyfile)
    content = content.replace("{name}", name)
    content = content.replace("{ttl_name}", ttl_name)
    content = content.replace("{logspath}", logspath)
    content = content.replace("{created_at}", timestamp)

    # グループ階層の取得
    group1 = str(row.get("group1", "") if pd.notna(row.get("group1")) else "").strip()
    group2 = str(row.get("group2", "") if pd.notna(row.get("group2")) else "").strip()
    group3 = str(row.get("group3", "") if pd.notna(row.get("group3")) else "").strip()

    # 有効な親階層がある場合のみ作成（子階層のみの指定は無効）
    if group1:
        target_dir = OUTPUT_DIR / group1
        if group2:
            target_dir = target_dir / group2
            if group3:
                target_dir = target_dir / group3
        target_dir.mkdir(parents=True, exist_ok=True)
    else:
        target_dir = OUTPUT_DIR

    (target_dir / f"{ttl_name}.ttl").write_text(content, encoding="utf-8")
    print(f"✅ {ttl_name}.ttl を生成しました。")
