from __future__ import annotations

"""Health Auto Export CSV/TSV → Sleep シートデータへの変換"""

import csv
import datetime
import io


# CSV列名 → (sleepキー, 変換関数)
CSV_TO_SLEEP = {
    # 心拍系
    "心拍数 [平均] (count/min)":          ("hearwatch.daily_bpm",                  float),
    "心拍数 [最小] (count/min)":          ("hearwatch.sleep_bpm",                  float),
    "安静時心拍数 (count/min)":           ("hearwatch.waking_bpm",                 float),
    "心拍変動 (ms)":                      ("hearwatch.sleep_hrv",                  float),
    "呼吸数 (count/min)":                 ("hearwatch.sleep_respiratory_rate",     float),
    # 睡眠ステージ（全ステージを取得）
    "睡眠分析 [Total] (hr)":              ("sleep.total_sleep",   lambda x: float(x) * 60),
    "睡眠分析 [深い] (hr)":               ("sleep.deep_sleep",    lambda x: float(x) * 60),
    "睡眠分析 [REM] (hr)":                ("sleep.rem_sleep",     lambda x: float(x) * 60),
    "睡眠分析 [コア] (hr)":               ("sleep.core_sleep",    lambda x: float(x) * 60),
    "睡眠分析 [起きている] (hr)":          ("sleep.awake_time",    lambda x: float(x) * 60),
    # 体温・体力
    "Apple 睡眠時手首温度 (degC)":         ("wellness.wrist_temp",                  float),
    "VO2 Max (ml/(kg·min))":              ("wellness.vo2_max",                     float),
}


def _detect_delimiter(text: str) -> str:
    """タブ区切りかカンマ区切りかを自動検出する。"""
    first_line = text.split("\n")[0] if "\n" in text else text
    return "\t" if "\t" in first_line else ","


def _parse_row(row: dict) -> dict:
    """
    CSVの1行からSleepシート書き込み用dictを生成する。
    Returns:
        {
          "hearwatch": { "daily_bpm": ..., ... },
          "sleep":     { "total_sleep": {"total_minutes": ...}, ... },
          "wellness":  { "wrist_temp": ..., ... },
        }
    """
    hearwatch: dict = {}
    sleep: dict = {}
    wellness: dict = {}

    for csv_col, (sleep_key, convert) in CSV_TO_SLEEP.items():
        raw = row.get(csv_col, "").strip()
        if not raw:
            continue
        try:
            val = convert(raw)
        except (ValueError, TypeError):
            continue
        if val is None:
            continue

        section, field = sleep_key.split(".", 1)
        if section == "hearwatch":
            hearwatch[field] = round(val, 2)
        elif section == "sleep":
            sleep[field] = {"total_minutes": round(val, 1)}
        elif section == "wellness":
            wellness[field] = round(val, 2)

    result = {}
    if hearwatch:
        result["hearwatch"] = hearwatch
    if sleep:
        result["sleep"] = sleep
    if wellness:
        result["wellness"] = wellness
    return result


def parse_file_content(csv_content: str) -> tuple[datetime.date | None, dict]:
    """
    1ファイル（1日分）の Health Export ファイルをパースする。
    Returns: (date, sleep_data_dict)
    """
    delim = _detect_delimiter(csv_content)
    reader = csv.DictReader(io.StringIO(csv_content), delimiter=delim)

    for row in reader:
        raw_date = row.get("日付/時間", "").strip()
        if not raw_date:
            continue
        try:
            row_date = datetime.datetime.strptime(raw_date, "%Y-%m-%d %H:%M:%S").date()
        except ValueError:
            try:
                row_date = datetime.datetime.strptime(raw_date, "%Y-%m-%d").date()
            except ValueError:
                continue

        data = _parse_row(row)
        return row_date, data

    return None, {}


def parse_health_csv(
    csv_content: str,
    target_date: datetime.date,
) -> dict:
    """
    複数行 CSV から特定日のデータを取り出す（後方互換）。
    Returns: {"hearwatch": {...}, "sleep": {...}, "_source_row": {...}}
    """
    delim = _detect_delimiter(csv_content)
    reader = csv.DictReader(io.StringIO(csv_content), delimiter=delim)
    target_row = None

    for row in reader:
        raw_date = row.get("日付/時間", "").strip()
        try:
            row_date = datetime.datetime.strptime(raw_date, "%Y-%m-%d %H:%M:%S").date()
        except ValueError:
            try:
                row_date = datetime.datetime.strptime(raw_date, "%Y-%m-%d").date()
            except ValueError:
                continue
        if row_date == target_date:
            target_row = row
            break

    if target_row is None:
        raise ValueError(f"{target_date} のデータがCSVに見つかりませんでした。")

    data = _parse_row(target_row)
    data["_source_row"] = dict(target_row)
    return data


def get_available_dates(csv_content: str) -> list[datetime.date]:
    """CSVに含まれる日付の一覧を返す。"""
    delim = _detect_delimiter(csv_content)
    reader = csv.DictReader(io.StringIO(csv_content), delimiter=delim)
    dates = []
    for row in reader:
        raw_date = row.get("日付/時間", "").strip()
        try:
            d = datetime.datetime.strptime(raw_date, "%Y-%m-%d %H:%M:%S").date()
            dates.append(d)
        except ValueError:
            try:
                d = datetime.datetime.strptime(raw_date, "%Y-%m-%d").date()
                dates.append(d)
            except ValueError:
                pass
    return sorted(dates)
