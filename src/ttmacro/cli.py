"""ttmacro CLI エントリポイント。

argparse でコマンドライン引数を解析し、Excel 台帳から TTL を生成する。

openpyxl に依存する excel_loader は ``generate_ttl_macros()`` 内で遅延 import
する。これにより ``--help`` や ``--clean --dry-run`` は openpyxl 未インストール
環境でも動作する。
"""

from __future__ import annotations

import argparse
import sys
import traceback
from datetime import datetime
from pathlib import Path

from ttmacro import path_resolver, ttl_renderer
from ttmacro.config import BASE_DIR, EXCEL_PATH
from ttmacro.logger import setup_logging


def parse_args() -> argparse.Namespace:
    """コマンドライン引数を解析する。

    Returns:
        ``--row`` / ``--clean`` / ``--dry-run`` を含む argparse.Namespace。

    Raises:
        SystemExit: ``--dry-run`` が ``--clean`` なしで指定された場合。
    """
    parser = argparse.ArgumentParser(
        description=r"""
TTLマクロを生成するツール

servers.xlsxの内容に基づいてTTLマクロを生成します。
generate列が'yes'の行のみが処理対象となります。
'e'が指定されている行で処理を終了します。

実行方法:
  # PowerShellの場合
  python .\generate_ttl_macros.py [オプション]

  # コマンドプロンプトの場合
  python generate_ttl_macros.py [オプション]
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=r"""
使用例:
  # 全行を生成
  python .\generate_ttl_macros.py

  # 特定の行のみ生成（5行目）
  python .\generate_ttl_macros.py --row 5

  # 既存 TTL を全削除してから生成（グループ変更時の孤児解消）
  python .\generate_ttl_macros.py --clean

  # 削除対象を確認するだけ（実削除も生成も行わない）
  python .\generate_ttl_macros.py --clean --dry-run

  # ヘルプを表示
  python .\generate_ttl_macros.py --help

注意:
  - 行番号はExcelのA列のNo.を指定します
  - generate列が'yes'の行のみが処理されます
  - 生成フラグに'e'を指定すると処理を終了します
  - PowerShellで実行する場合は 'python .\generate_ttl_macros.py' を使用してください
  - --row と --clean は併用できません
        """,
    )

    # --row と --clean は排他（--row は単一行処理、--clean は全体クリーン+全行生成）
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--row",
        type=int,
        help="生成する行番号（1から始まる）。指定がない場合は全行を処理します。",
    )
    group.add_argument(
        "--clean",
        action="store_true",
        help="生成前に既存の TTL ファイルを全削除する（template.ttl 除く）。"
        "グループ変更で残った旧 TTL を掃除する用途。",
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="--clean と組み合わせて、削除対象を表示するだけで実際の削除と生成は行わない。",
    )

    args = parser.parse_args()

    # --dry-run は --clean とのみ併用可（単独使用は意味がない）
    if args.dry_run and not args.clean:
        parser.error("--dry-run は --clean とのみ併用できます")

    return args


def generate_ttl_macros(args: argparse.Namespace) -> None:
    """TTLマクロを生成するメイン関数。

    Excel 台帳の各行を ``generate=yes`` フィルタで処理し、
    テンプレートを展開して TTL ファイルを書き出す。``'e'`` を検出すると終了。

    ``--clean`` 指定時は生成前に既存 TTL を一括削除する。``--dry-run`` を
    併用すると削除対象を出力するだけで実削除・生成ともにスキップする。

    Args:
        args: ``parse_args()`` の戻り値。
    """
    print("[1/4] ログ設定...", file=sys.stderr, flush=True)
    logger = setup_logging()

    try:
        # クリーンアップフェーズ（--clean 指定時のみ）
        if args.clean:
            from ttmacro import cleaner
            from ttmacro.config import OUTPUT_DIR, TEMPLATES_DIR

            # TEMPLATES_DIR 配下を保護（テンプレートを誤削除しない）
            targets = cleaner.find_ttl_files_to_delete(OUTPUT_DIR, TEMPLATES_DIR)
            existing_empty = cleaner.find_empty_subdirs(OUTPUT_DIR)
            logger.info(
                f"🧹 クリーン対象: TTL {len(targets)} 件、"
                f"既存の空ディレクトリ {len(existing_empty)} 件"
            )

            if args.dry_run:
                logger.info("🔍 [dry-run] 削除候補一覧:")
                for t in targets:
                    logger.info(f"  - {t.relative_to(OUTPUT_DIR)}")
                logger.info(
                    "🔍 [dry-run] 実削除と生成は行いません（--dry-run なしで実行してください）"
                )
                return

            file_count = cleaner.delete_ttl_files(targets)
            dir_count = cleaner.delete_empty_subdirs(OUTPUT_DIR)
            logger.info(
                f"🧹 クリーン完了: TTL {file_count} 件、"
                f"空ディレクトリ {dir_count} 件を削除"
            )

        # openpyxl に依存する excel_loader はここで遅延 import
        # （--help と --clean --dry-run を openpyxl 未導入環境でも動かすため）
        try:
            from ttmacro import excel_loader
        except ImportError as e:
            print(
                f"openpyxl のインポートに失敗しました: {e}",
                file=sys.stderr,
                flush=True,
            )
            print(
                '仮想環境を有効にして、pip install -e ".[dev]" を実行してください。',
                file=sys.stderr,
                flush=True,
            )
            sys.exit(1)

        print("[2/4] テンプレート・Excel 読み込み...", file=sys.stderr, flush=True)
        # 行ごとに異なるテンプレを許容するため、ここでは Excel のみ先読み。
        # テンプレは行処理ループでパス解決して読み込む（同一パスはキャッシュ）。
        template_cache: dict[Path, str] = {}
        headers, rows = excel_loader.load_excel_data()
        timestamp = datetime.now().strftime("%Y/%m/%d %H:%M:%S")

        logger.info(f"読み込み元: {EXCEL_PATH}")
        logger.info("生成開始")

        # 必要な列の存在チェック
        required_columns = ["No.", "name", "host", "user", "generate"]
        missing_columns = [col for col in required_columns if col not in headers]
        if missing_columns:
            raise ValueError(f"必要な列が見つかりません: {', '.join(missing_columns)}")

        # 行番号が指定されている場合
        if args.row is not None:
            matching = [
                (idx, r) for idx, r in enumerate(rows) if r.get("No.") == args.row
            ]
            if not matching:
                logger.error(f"❌ 指定されたNo. {args.row} は見つかりませんでした")
                return
            rows_to_process = [matching[0]]
            logger.info(f"📝 No.{args.row} のサーバーを処理します")
        else:
            rows_to_process = list(enumerate(rows))
            # generate=yes の行数を事前に表示
            generate_count = sum(
                1
                for r in rows
                if str(r.get("generate", "") or "").strip().lower()
                in ("yes", "true", "1")
            )
            logger.info(
                f"generate=yes の行: {generate_count} 件（全 {len(rows)} 行中）"
            )
            if generate_count == 0:
                logger.warning(
                    "⚠️ 対象行が0件です。Excelの generate 列に yes を指定した行がありますか？"
                )

        print("[3/4] 行を処理しています...", file=sys.stderr, flush=True)
        success_count = 0
        error_count = 0

        for idx, row in rows_to_process:
            try:
                # 空白行はスキップ
                if excel_loader.is_blank_row(row):
                    continue

                # 生成フラグの確認（--row 指定時はフラグ無視）
                generate_flag = str(row.get("generate", "") or "").strip().lower()
                if args.row is None and generate_flag == "e":
                    logger.info("⏹️ 'e' を検出したため、処理を終了します。")
                    break
                # yes/true/1 を有効とする（Excel の TRUE や 1 にも対応）
                if args.row is None and generate_flag not in ("yes", "true", "1"):
                    continue

                # 行データの検証（No. 列が int でない/空ならインデックス番号を流用）
                no_value = row.get("No.")
                try:
                    row_num = (
                        idx + 1 if excel_loader.is_blank(no_value) else int(no_value)  # type: ignore[arg-type]
                    )
                except (TypeError, ValueError):
                    row_num = idx + 1
                is_valid, validation_errors = excel_loader.validate_row_data(
                    row, row_num
                )
                if not is_valid:
                    error_msg = (
                        f"No.{row_num} データ検証エラー: {'; '.join(validation_errors)}"
                    )
                    logger.error(f"❌ {error_msg}")
                    error_count += 1
                    continue

                # 行データの抽出と TTL 生成
                data = excel_loader.extract_row_data(row)
                target_dir = path_resolver.get_target_directory(data)

                # テンプレ解決＋キャッシュ。失敗はその行だけスキップ。
                try:
                    template_path = ttl_renderer.resolve_template_path(data["template"])
                    if template_path not in template_cache:
                        template_cache[template_path] = ttl_renderer.load_template(
                            template_path
                        )
                    template = template_cache[template_path]
                except (FileNotFoundError, ValueError, RuntimeError) as e:
                    logger.error(
                        f"❌ No.{row_num} テンプレート読込エラー "
                        f"(template='{data['template']}'): {e}"
                    )
                    error_count += 1
                    continue

                content = ttl_renderer.generate_ttl_content(
                    data, template, timestamp, target_dir
                )

                ttl_name = f"{data['name']}_{data['host']}_{data['user']}"
                ttl_file = target_dir / f"{ttl_name}.ttl"

                try:
                    ttl_file.write_text(content, encoding="utf-8")
                    logger.info(f"✅ {ttl_name}.ttl を生成しました。（No.{row_num}）")
                    success_count += 1
                except Exception as e:
                    logger.error(f"❌ ファイル書き込みエラー {ttl_name}.ttl: {e}")
                    error_count += 1

            except Exception as e:
                # 例外時のログ用 No.（int 化失敗時はインデックス番号で代替）
                no_value = row.get("No.")
                try:
                    fallback_num: int = (
                        idx + 1
                        if excel_loader.is_blank_row(row)
                        or excel_loader.is_blank(no_value)
                        else int(no_value)  # type: ignore[arg-type]
                    )
                except (TypeError, ValueError):
                    fallback_num = idx + 1
                logger.error(f"❌ No.{fallback_num} 処理エラー: {e}")
                error_count += 1

        print("[4/4] 完了", file=sys.stderr, flush=True)
        logger.info(f"📊 処理完了 - 成功: {success_count}件, エラー: {error_count}件")

    except Exception as e:
        err_msg = f"致命的エラー: {e}"
        tb_lines = traceback.format_exc()
        # ログ未初期化でも確実に表示するため stderr に直接出す
        print("", file=sys.stderr)
        print("=== エラー内容（トレースバック） ===", file=sys.stderr)
        print(tb_lines, file=sys.stderr)
        print("====================================", file=sys.stderr)
        try:
            logger.error(f"❌ {err_msg}")
            logger.error("トレースバック:\n%s", tb_lines)
        except NameError:
            print(f"エラー: {err_msg}", file=sys.stderr)
        sys.stderr.flush()
        sys.exit(1)


def main() -> None:
    """ttmacro CLI のエントリポイント。"""
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


if __name__ == "__main__":
    main()
