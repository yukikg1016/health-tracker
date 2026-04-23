from __future__ import annotations

"""Health Auto Export CSV → Sleep シートデータへの変換"""

import csv
import datetime
import io


# CSV列名 → (sleepキー, 変換関数)
CSV_TO_SLEEP = {
    "心拍数 [平均] (count/min)":      ("hearwatch.daily_bpm",                  float),
    "呼吸数 (count/min)":             ("hearwatch.sleep_respiratory_rate",      float),
    "心拍変動 (ms)":                  ("hearwatch.sleep_hrv",                   float),
    "安静時心拍数 (count/min)":       ("hearwatch.waking_bpm",                  float),
    "睡眠分析 [Total] (hr)":          ("sleep.total_sleep",                     lambda x: float(x) * 60),
    "睡眠分析 [深い] (hr)":           ("sleep.deep_sleep",                      lambda x: float(x) * 60),
    "心拍数 [最小] (count/min)":      ("hearwatch.sleep_bpm",                   float),
}


def parse_health_csv(
    csv_content: str,
    target_date: datetime.date,
) -> dict:
    """
    Health Auto Export CSVを解析してSleepシート書き込み用dictを返す。

    Returns:
        {
          "hearwatch": { "daily_bpm": ..., "sleep_hrv": ..., ... },
          "sleep":     { "total_sleep": {"total_minutes": ...}, ... },
          "_source_row": {...}   # 元データ（プレビュー用）
        }
    """
    reader = csv.DictReader(io.StringIO(csv_content), delimiter="\t")
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

    hearwatch: dict = {}
    sleep: dict = {}

    for csv_col, (sleep_key, convert) in CSV_TO_SLEEP.items():
        raw = target_row.get(csv_col, "").strip()
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
            # sleep_writer は duration を {"total_minutes": x} の形式で受け取る
            sleep[field] = {"total_minutes": round(val, 1)}

    return {
        "hearwatch": hearwatch,
        "sleep":     sleep,
        "_source_row": dict(target_row),
    }


def get_available_dates(csv_content: str) -> list[datetime.date]:
    """CSVに含まれる日付の一覧を返す。"""
    reader = csv.DictReader(io.StringIO(csv_content), delimiter="\t")
    dates = []
    for row in reader:
        raw_date = row.get("日付/時間", "").strip()
        try:
            d = datetime.datetime.strptime(raw_date, "%Y-%m-%d %H:%M:%S").date()
            dates.append(d)
        except ValueError:
            pass
    return sorted(dates)
