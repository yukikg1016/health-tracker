from __future__ import annotations

"""SleepシートへのExcel書き込み"""

import datetime
from .writer_base import ExcelWriter

# Sleepシートの行マップ（Excel行番号 1-based、実際のシートを確認済み）
SLEEP_ROW_MAP = {
    # HeartWatch セクション
    "hearwatch.daily_bpm":              7,
    "hearwatch.sedentary_bpm":          8,
    "hearwatch.sleep_bpm":              9,
    "hearwatch.sleep_spo2":             10,
    "hearwatch.sleep_respiratory_rate": 11,
    "hearwatch.sleep_time":             12,   # total_minutes を書き込む
    "hearwatch.restfulness":            13,
    "hearwatch.sleep_hrv":              14,
    "hearwatch.waking_bpm":             15,
    "hearwatch.waking_hrv":             16,
    "hearwatch.daily_spo2":             17,
    # Sleep セクション
    "sleep.total_sleep":                21,   # total_minutes を書き込む
    "sleep.sleep_quality":              22,
    "sleep.deep_sleep":                 23,   # total_minutes を書き込む
    "sleep.bpm_sleep_avg":              24,
    "sleep.sleep_session_start":        25,
    "sleep.sleep_session_end":          26,
    "sleep.sleep_efficiency":           27,
    "sleep.sleep_rating":               28,
    # Wellness セクション
    "wellness.readiness_score":         31,
    "wellness.sleep_fuel_rating":       32,
    "wellness.hrv":                     33,
    "wellness.bpm_hrv":                 34,
    "wellness.prior_day_stress":        35,
    "wellness.stress_hrv":              36,
    "wellness.sleep_bank":              37,
    "wellness.sleep_bank_balance":      38,
    "wellness.wrist_temp":              40,
    "wellness.wrist_temp_baseline":     41,
    "wellness.wrist_temp_deviation":    42,
    "wellness.sleep_spo2_avg":          44,
    "wellness.sleep_spo2_range":        45,
    "wellness.respiration_rate_avg":    47,
    "wellness.respiration_rate_range":  48,
}

# 時間系フィールド（total_minutes で書く）
DURATION_KEYS = {
    "hearwatch.sleep_time",
    "sleep.total_sleep",
    "sleep.deep_sleep",
}


def build_sleep_preview(data: dict, col_idx: int) -> list[dict]:
    """プレビュー用の行リストを返す。"""
    rows = []
    for key, row_idx in SLEEP_ROW_MAP.items():
        section, field = key.split(".", 1)
        section_data = data.get(section, {})

        if key in DURATION_KEYS:
            val = (section_data.get(field) or {}).get("total_minutes")
        else:
            val = section_data.get(field)

        rows.append({
            "フィールド": key,
            "抽出値":     val,
            "書き込み先": f"行{row_idx} 列{col_idx}",
        })
    return rows


def write_sleep_data(
    data: dict,
    target_date: datetime.date,
    excel_path: str,
    skip_existing: bool = False,
) -> list[dict]:
    """
    Sleepシートに書き込む。

    Args:
        skip_existing: True の場合、既に値があるセルはスキップする（Health CSV インポート用）

    Returns:
        プレビュー行リスト
    """
    writer = ExcelWriter(excel_path)
    wb = writer.load()
    ws = wb[ExcelWriter.SHEET_SLEEP]

    col_idx = ExcelWriter.get_or_create_date_column(ws, target_date)
    preview = build_sleep_preview(data, col_idx)

    for key, row_idx in SLEEP_ROW_MAP.items():
        section, field = key.split(".", 1)
        section_data = data.get(section) or {}

        if key in DURATION_KEYS:
            val = (section_data.get(field) or {}).get("total_minutes")
        else:
            val = section_data.get(field)

        if val is not None:
            # dict が来た場合は total_minutes → hours 順に変換
            if isinstance(val, dict):
                if "total_minutes" in val:
                    val = val["total_minutes"]
                elif "hours" in val and "minutes" in val:
                    val = val["hours"] * 60 + val["minutes"]
                else:
                    continue  # 変換できない辞書はスキップ

            cell = ws.cell(row=row_idx, column=col_idx)
            # skip_existing=True の場合、既存値があればスキップ
            if skip_existing and cell.value is not None:
                continue
            # セルがパーセント書式の場合は0〜1の小数に変換して書き込む
            fmt = cell.number_format or ""
            if "%" in fmt and isinstance(val, (int, float)):
                val = val / 100.0
            cell.value = val

    writer.save(wb)
    return preview
