from __future__ import annotations

"""夜💤シートへのExcel書き込み"""

import datetime
from .writer_base import ExcelWriter

# 夜シートの列マップ（実際のヘッダー確認済み）
NIGHT_COL_MAP = {
    "timestamp":        1,   # タイムスタンプ（新規行のみ自動設定）
    "date":             2,   # 日付
    "water_intake_ml":  3,   # 水分摂取量
    # col 4: 食事内容（テキスト）— スクリーンショットからは抽出しない
    # col 5: 食事写真URL — スクリーンショットからは抽出しない
    "training_rpe":     6,   # 練習強度
    "study_hours":      7,   # 勉強時間
    "jump_cm":          8,   # ジャンプ記録
    "notes":            9,   # 追記事項
    "coffee_cups":      10,  # コーヒー
    "run_distance_km":  11,  # ランニング距離
    # col 12: 列11（スペア）
    "run_power_w":      13,  # ランニングパワー
    "pushups":          14,  # プッシュアップ
    "grip_strength_kg": 15,  # 握力
    # col 16: 列8（スペア）
}

DATE_COL = 2


def _find_or_create_row(ws, target_date: datetime.date) -> tuple[int, bool]:
    for row in range(2, ws.max_row + 2):
        val = ws.cell(row=row, column=DATE_COL).value
        if val is None:
            return row, True
        if isinstance(val, datetime.datetime):
            val = val.date()
        if isinstance(val, datetime.date) and val == target_date:
            return row, False
    return ws.max_row + 1, True


def build_night_preview(data: dict, row_idx: int, is_new: bool) -> list[dict]:
    rows = []
    for key, col in NIGHT_COL_MAP.items():
        if key in ("timestamp", "date"):
            continue
        rows.append({
            "フィールド": key,
            "抽出値":     data.get(key),
            "書き込み先": f"行{row_idx} 列{col}",
            "新規行":     "✓" if is_new else "上書き",
        })
    return rows


def write_night_data(data: dict, target_date: datetime.date, excel_path: str) -> list[dict]:
    """
    夜💤シートに書き込む。同日の行があれば上書き、なければ追加。

    Returns:
        プレビュー行リスト
    """
    writer = ExcelWriter(excel_path)
    wb = writer.load()
    ws = wb[ExcelWriter.SHEET_NIGHT]

    row_idx, is_new = _find_or_create_row(ws, target_date)
    preview = build_night_preview(data, row_idx, is_new)

    if is_new:
        ws.cell(row=row_idx, column=1).value = datetime.datetime.now()
        ws.cell(row=row_idx, column=2).value = datetime.datetime(
            target_date.year, target_date.month, target_date.day
        )

    for key, col in NIGHT_COL_MAP.items():
        if key in ("timestamp", "date"):
            continue
        val = data.get(key)
        if val is not None:
            ws.cell(row=row_idx, column=col).value = val

    writer.save(wb)
    return preview
