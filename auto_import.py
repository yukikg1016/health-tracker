"""
Auto Import — GitHub Actions から毎朝実行される自動インポートスクリプト。
Google Drive の HealthMetrics フォルダをスキャンし、
未書き込みの睡眠データを Excel に書き込んで Drive に戻す。

使い方:
    python auto_import.py [--days 7] [--dry-run]

必要な環境変数:
    GOOGLE_SERVICE_ACCOUNT_JSON   (サービスアカウントJSONの文字列)
    GDRIVE_FILE_ID                (ExcelファイルのDrive ID)
    USER_1_HEALTH_FOLDER_ID       (ユーザー1のHealthExportフォルダID)
    USER_2_HEALTH_FOLDER_ID       (ユーザー2、省略可)
    USER_3_HEALTH_FOLDER_ID       (省略可)
    USER_4_HEALTH_FOLDER_ID       (省略可)
"""

from __future__ import annotations

import argparse
import datetime
import os
import pathlib
import sys

# プロジェクトルートをパスに追加
BASE_DIR = pathlib.Path(__file__).parent
sys.path.insert(0, str(BASE_DIR))


def _get_user_configs() -> list[dict]:
    """設定済みユーザーの一覧を返す。"""
    configs = []
    for uid in ["1", "2", "3", "4"]:
        folder_id = os.environ.get(f"USER_{uid}_HEALTH_FOLDER_ID", "").strip()
        if folder_id:
            # User 1 はシートプレフィックスなし、2以降は "2_" のプレフィックス
            prefix = "" if uid == "1" else f"{uid}_"
            configs.append({
                "user_id":      uid,
                "folder_id":    folder_id,
                "sheet_prefix": prefix,
            })
    return configs


def run_import(days: int = 7, dry_run: bool = False) -> None:
    from gdrive_helper import (
        download_to_temp,
        upload_from_path,
        list_health_export_files,
        download_file_as_text,
    )
    from extractors.health_csv_extractor import parse_file_content
    from excel_writer.sleep_writer import write_sleep_data

    users = _get_user_configs()
    if not users:
        print("⚠️  インポート対象ユーザーなし（USER_N_HEALTH_FOLDER_ID 未設定）")
        return

    # 対象期間（今日から days 日前まで）
    today = datetime.date.today()
    cutoff = today - datetime.timedelta(days=days)
    print(f"📅 対象期間: {cutoff} 〜 {today}")

    # Excel をダウンロード（全ユーザー共通ファイル）
    print("⬇️  Excel ダウンロード中...")
    excel_path = download_to_temp()
    print(f"   → {excel_path}")

    total_written = 0
    total_skipped = 0
    total_errors  = 0

    for user in users:
        uid        = user["user_id"]
        folder_id  = user["folder_id"]
        prefix     = user["sheet_prefix"]
        print(f"\n👤 User {uid}  (prefix='{prefix}', folder={folder_id[:12]}...)")

        try:
            files = list_health_export_files(folder_id)
        except Exception as e:
            print(f"   ❌ フォルダ取得エラー: {e}")
            total_errors += 1
            continue

        health_files = [
            f for f in files
            if "HealthMetrics" in f.get("name", "")
            or "health" in f.get("name", "").lower()
        ]
        print(f"   📁 {len(health_files)} ファイル検出")

        for f in health_files:
            try:
                content = download_file_as_text(f["id"], f.get("mimeType", ""))
                date, data = parse_file_content(content)
            except Exception as e:
                print(f"   ⚠️  {f['name']}: パースエラー {e}")
                total_errors += 1
                continue

            if date is None or not data:
                print(f"   —  {f['name']}: データなし")
                total_skipped += 1
                continue

            if date < cutoff:
                print(f"   —  {date}: 対象期間外（スキップ）")
                total_skipped += 1
                continue

            if dry_run:
                print(f"   [DRY-RUN] {date}: {list(data.keys())}")
                total_written += 1
                continue

            try:
                write_sleep_data(
                    data, date, excel_path,
                    skip_existing=True,
                    sheet_prefix=prefix,
                )
                print(f"   ✅ {date}: 書き込み完了")
                total_written += 1
            except Exception as e:
                print(f"   ❌ {date}: 書き込みエラー {e}")
                total_errors += 1

    # Excel をアップロード
    if not dry_run and total_written > 0:
        print(f"\n⬆️  Excel アップロード中...")
        upload_from_path(excel_path)
        print("   → アップロード完了")

    pathlib.Path(excel_path).unlink(missing_ok=True)

    print(f"\n{'='*40}")
    print(f"完了: 書き込み {total_written} 件 / スキップ {total_skipped} 件 / エラー {total_errors} 件")

    if total_errors > 0:
        sys.exit(1)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Auto Health Export importer")
    parser.add_argument("--days",    type=int,  default=7,     help="何日前まで遡るか（デフォルト7）")
    parser.add_argument("--dry-run", action="store_true",      help="実際には書き込まない")
    args = parser.parse_args()

    run_import(days=args.days, dry_run=args.dry_run)
