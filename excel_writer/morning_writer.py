from __future__ import annotations

"""朝☀️シートへのExcel書き込み"""

import datetime
from .writer_base import ExcelWriter

# 朝シートの列マップ（実際のヘッダー確認済み）
MORNING_COL_MAP = {
    "timestamp":         1,   # タイムスタンプ（新規行のみ自動設定）
    "user_id":           2,   # ユーザーID（固定値 1）
    "date":              3,   # 記録日
    "reaction_time_sec": 4,   # 停止信号テスト反応時間
    "alcohol_drinks":    5,   # アルコール摂取量
    "coffee_after4pm":   6,   # コーヒー（16時以降）
    "fatigue_1to10":     7,   # 疲労感
    "stress_1to10":      8,   # ストレス
    "device_before_bed": 9,   # デバイス使用
    "mood_1to10":        10,  # 気分
    "accuracy_pct":      11,  # 精度
    "sleep_duration_text": 12,
    "deep_sleep_text":   13,
    "sleep_score":       14,
    "test_result_s":     15,
}

DATE_COL = 3


def _find_or_create_row(ws, target_date: datetime.date) -> tuple[int, bool]:
    """
    日付列を走査して対象日の行を探す。
    Returns: (row_idx, is_new_row)
    """
    for row in range(2, ws.max_row + 2):
        val = ws.cell(row=row, column=DATE_COL).value
        if val is None:
            return row, True
        if isinstance(val, datetime.datetime):
            val = val.date()
        if isinstance(val, datetime.date) and val == target_date:
            return row, False
    return ws.max_row + 1, True


def build_morning_preview(data: dict, row_idx: int, is_new: bool) -> list[dict]:
    rows = []
    for key, col in MORNING_COL_MAP.items():
        if key in ("timestamp", "user_id", "date"):
            continue
        rows.append({
            "フィールド": key,
            "抽出値":     data.get(key),
            "書き込み先": f"行{row_idx} 列{col}",
            "新規行":     "✓" if is_new else "上書き",
        })
    return rows


def write_morning_data(data: dict, target_date: datetime.date, excel_path: str) -> list[dict]:
    """
    朝☀️シートに書き込む。同日の行があれば上書き、なければ追加。

    Returns:
        プレビュー行リスト
    """
    writer = ExcelWriter(excel_path)
    wb = writer.load()
    ws = wb[ExcelWriter.SHEET_MORNING]

    row_idx, is_new = _find_or_create_row(ws, target_date)
    preview = build_morning_preview(data, row_idx, is_new)

    if is_new:
        ws.cell(row=row_idx, column=1).value = datetime.datetime.now()
        ws.cell(row=row_idx, column=2).value = 1
        ws.cell(row=row_idx, column=3).value = datetime.datetime(
            target_date.year, target_date.month, target_date.day
        )

    for key, col in MORNING_COL_MAP.items():
        if key in ("timestamp", "user_id", "date"):
            continue
        val = data.get(key)
        if val is not None:
            ws.cell(row=row_idx, column=col).value = val

    writer.save(wb)
    return preview
