from __future__ import annotations

"""In BodyシートへのExcel書き込み"""

import datetime
from .writer_base import ExcelWriter

INBODY_HEADERS = [
    "Date", "Weight", "BMI", "Body Fat %", "Muscle Mass",
    "Body Fat Mass", "Visceral Fat Level", "Total Body Water",
    "Bone Mass", "Basal Metabolic Rate", "Skeletal Muscle Mass",
]

INBODY_COL_MAP = {
    "date":                 1,
    "weight":               2,
    "bmi":                  3,
    "body_fat_pct":         4,
    "muscle_mass":          5,
    "body_fat_mass":        6,
    "visceral_fat_level":   7,
    "total_body_water":     8,
    "bone_mass":            9,
    "basal_metabolic_rate": 10,
    "skeletal_muscle_mass": 11,
}


def _initialize_if_empty(ws) -> None:
    """シートが空なら（row 1 col 1 が None）ヘッダー行を作成する。"""
    if ws.cell(row=1, column=1).value is None:
        for col, header in enumerate(INBODY_HEADERS, start=1):
            ws.cell(row=1, column=col).value = header


def _find_or_create_row(ws, target_date: datetime.date) -> int:
    """Date列（col 1）を走査し、一致する行を返す。なければ末尾に追加。"""
    for row in range(2, ws.max_row + 2):
        val = ws.cell(row=row, column=1).value
        if val is None:
            return row
        if isinstance(val, datetime.datetime):
            val = val.date()
        if isinstance(val, datetime.date) and val == target_date:
            return row
    return ws.max_row + 1


def build_inbody_preview(data: dict, row_idx: int) -> list[dict]:
    rows = []
    for key, col in INBODY_COL_MAP.items():
        if key == "date":
            continue
        rows.append({
            "フィールド": key,
            "抽出値":     data.get(key),
            "書き込み先": f"行{row_idx} 列{col}",
        })
    return rows


def write_inbody_data(data: dict, target_date: datetime.date, excel_path: str) -> list[dict]:
    """
    In Bodyシートに書き込む。初回はヘッダーを自動作成。

    Returns:
        プレビュー行リスト
    """
    writer = ExcelWriter(excel_path)
    wb = writer.load()
    ws = wb[ExcelWriter.SHEET_INBODY]

    _initialize_if_empty(ws)
    row_idx = _find_or_create_row(ws, target_date)
    preview = build_inbody_preview(data, row_idx)

    # 日付を書き込む
    ws.cell(row=row_idx, column=1).value = datetime.datetime(
        target_date.year, target_date.month, target_date.day
    )

    for key, col in INBODY_COL_MAP.items():
        if key == "date":
            continue
        val = data.get(key)
        if val is not None:
            ws.cell(row=row_idx, column=col).value = val

    writer.save(wb)
    return preview
