"""
Workout writer — appends one row per workout to the "Workout" sheet.
Sheet structure (auto-created if missing):
  Row 1: headers
  Row 2+: data rows
"""
import datetime
import openpyxl

HEADERS = [
    "Date",
    "Type",
    "Workout Time (min)",
    "Elapsed Time (min)",
    "Distance (km)",
    "Active kcal",
    "Total kcal",
    "Avg Heart Rate (bpm)",
    "Avg Pace (min/km)",
    "Avg Power (W)",
    "Avg Cadence (spm)",
    "Effort",
    "Move (kcal)",
    "Move Goal (kcal)",
    "Exercise (min)",
    "Exercise Goal (min)",
    "Stand (h)",
    "Stand Goal (h)",
    "Step Count",
    "Step Distance (km)",
]

SHEET_NAME = "Workout"


def _ensure_sheet(wb):
    if SHEET_NAME not in wb.sheetnames:
        ws = wb.create_sheet(SHEET_NAME)
        for col, h in enumerate(HEADERS, 1):
            ws.cell(row=1, column=col, value=h)
        return ws
    return wb[SHEET_NAME]


def write_workout_data(data: dict, target_date, excel_path: str) -> None:
    wb = openpyxl.load_workbook(excel_path)
    ws = _ensure_sheet(wb)

    if isinstance(target_date, str):
        target_date = datetime.date.fromisoformat(target_date)

    row = [
        target_date,
        data.get("workout_type"),
        data.get("workout_time_min"),
        data.get("elapsed_time_min"),
        data.get("distance_km"),
        data.get("active_kcal"),
        data.get("total_kcal"),
        data.get("avg_heart_rate_bpm"),
        data.get("avg_pace_min_per_km"),
        data.get("avg_power_watts"),
        data.get("avg_cadence_spm"),
        data.get("effort"),
        data.get("move_kcal"),
        data.get("move_goal_kcal"),
        data.get("exercise_min"),
        data.get("exercise_goal_min"),
        data.get("stand_hours"),
        data.get("stand_goal_hours"),
        data.get("step_count"),
        data.get("step_distance_km"),
    ]

    ws.append(row)
    wb.save(excel_path)


def build_workout_preview(data: dict, row_idx: int) -> list:
    rows = []
    field_map = [
        ("workout_type",        "Type",                  ""),
        ("workout_time_min",    "Workout Time (min)",    ""),
        ("elapsed_time_min",    "Elapsed Time (min)",    ""),
        ("distance_km",         "Distance (km)",         ""),
        ("active_kcal",         "Active kcal",           ""),
        ("total_kcal",          "Total kcal",            ""),
        ("avg_heart_rate_bpm",  "Avg Heart Rate (bpm)",  ""),
        ("avg_pace_min_per_km", "Avg Pace (min/km)",     ""),
        ("avg_power_watts",     "Avg Power (W)",         ""),
        ("avg_cadence_spm",     "Avg Cadence (spm)",     ""),
        ("effort",              "Effort",                ""),
        ("move_kcal",           "Move (kcal)",           "Activity Rings"),
        ("move_goal_kcal",      "Move Goal (kcal)",      "Activity Rings"),
        ("exercise_min",        "Exercise (min)",        "Activity Rings"),
        ("exercise_goal_min",   "Exercise Goal (min)",   "Activity Rings"),
        ("stand_hours",         "Stand (h)",             "Activity Rings"),
        ("stand_goal_hours",    "Stand Goal (h)",        "Activity Rings"),
        ("step_count",          "Step Count",            ""),
        ("step_distance_km",    "Step Distance (km)",    ""),
    ]
    for key, label, category in field_map:
        val = data.get(key)
        if val is not None:
            rows.append({"項目": label, "値": val, "カテゴリ": category, "書き込み先": f"行{row_idx}"})
    return rows
