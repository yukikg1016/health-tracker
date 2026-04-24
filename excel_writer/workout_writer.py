"""
Workout writer — writes to fixed rows per workout type, columns = dates.
Sheet structure matches the existing Workout sheet layout.
"""
import datetime
import openpyxl
from .writer_base import ExcelWriter

SHEET_NAME = "Workout"

# Row map per workout type: field → row number
CORE_TRAINING_ROWS = {
    "workout_time_min":   4,
    "elapsed_time_min":   5,
    "active_kcal":        6,
    "total_kcal":         7,
    "avg_heart_rate_bpm": 8,
    "effort":             9,
}

OUTDOOR_RUN_ROWS = {
    "workout_time_min":   13,
    "distance_km":        14,
    "active_kcal":        15,
    "total_kcal":         16,
    "avg_power_watts":    17,
    "avg_cadence_spm":    18,
    "avg_pace_min_per_km": 19,
    "avg_heart_rate_bpm": 20,
    "effort":             21,
}

OUTDOOR_WALK_ROWS = {
    "workout_time_min":   25,
    "distance_km":        26,
    "active_kcal":        27,
    "total_kcal":         28,
    "avg_pace_min_per_km": 29,
    "avg_heart_rate_bpm": 30,
    "effort":             31,
}

ACTIVITY_RINGS_ROWS = {
    "move_kcal":        36,
    "exercise_min":     37,
    "stand_hours":      38,
    "step_count":       39,
    "step_distance_km": 40,
}

# workout_type string → row map
TYPE_ROW_MAP = {
    "core training":   CORE_TRAINING_ROWS,
    "outdoor run":     OUTDOOR_RUN_ROWS,
    "outdoor walk":    OUTDOOR_WALK_ROWS,
    "activity rings":  ACTIVITY_RINGS_ROWS,
    "strength training": CORE_TRAINING_ROWS,  # fallback to core rows
}


def _get_row_map(workout_type: str) -> dict:
    if not workout_type:
        return {}
    return TYPE_ROW_MAP.get(workout_type.lower().strip(), CORE_TRAINING_ROWS)


def write_workout_data(data: dict, target_date, excel_path: str) -> None:
    if isinstance(target_date, str):
        target_date = datetime.date.fromisoformat(target_date)

    writer = ExcelWriter(excel_path)
    wb = writer.load()
    ws = wb[SHEET_NAME]

    # Find or create date column
    col_idx = ExcelWriter.find_date_column(ws, target_date)
    if col_idx is None:
        col_idx = ws.max_column + 1
        ws.cell(row=1, column=col_idx).value = target_date.strftime("%Y.%m.%d")

    workout_type = data.get("workout_type", "")
    row_map = _get_row_map(workout_type)

    for field, row in row_map.items():
        val = data.get(field)
        if val is not None:
            ws.cell(row=row, column=col_idx).value = val

    # Also write move_goal and exercise_goal if Activity Rings
    if workout_type and "activity" in workout_type.lower():
        # Write "current/goal" format in Move cell
        move = data.get("move_kcal")
        move_goal = data.get("move_goal_kcal")
        if move is not None and move_goal is not None:
            ws.cell(row=36, column=col_idx).value = f"{move}/{move_goal}"
        elif move is not None:
            ws.cell(row=36, column=col_idx).value = move

        ex = data.get("exercise_min")
        ex_goal = data.get("exercise_goal_min")
        if ex is not None and ex_goal is not None:
            ws.cell(row=37, column=col_idx).value = f"{ex}/{ex_goal}"
        elif ex is not None:
            ws.cell(row=37, column=col_idx).value = ex

        st = data.get("stand_hours")
        st_goal = data.get("stand_goal_hours")
        if st is not None and st_goal is not None:
            ws.cell(row=38, column=col_idx).value = f"{st}/{st_goal}"
        elif st is not None:
            ws.cell(row=38, column=col_idx).value = st

    writer.save(wb)


def build_workout_preview(data: dict, row_idx: int) -> list:
    workout_type = data.get("workout_type", "")
    row_map = _get_row_map(workout_type)

    label_names = {
        "workout_time_min":    "Workout Time (min)",
        "elapsed_time_min":    "Elapsed Time (min)",
        "distance_km":         "Distance (km)",
        "active_kcal":         "Active kcal",
        "total_kcal":          "Total kcal",
        "avg_heart_rate_bpm":  "Avg Heart Rate (bpm)",
        "avg_pace_min_per_km": "Avg Pace (min/km)",
        "avg_power_watts":     "Avg Power (W)",
        "avg_cadence_spm":     "Avg Cadence (spm)",
        "effort":              "Effort",
        "move_kcal":           "Move (kcal)",
        "move_goal_kcal":      "Move Goal (kcal)",
        "exercise_min":        "Exercise (min)",
        "exercise_goal_min":   "Exercise Goal (min)",
        "stand_hours":         "Stand (h)",
        "stand_goal_hours":    "Stand Goal (h)",
        "step_count":          "Step Count",
        "step_distance_km":    "Step Distance (km)",
    }

    rows = [{"項目": "Type", "値": workout_type, "行": "-"}]
    for field, row in row_map.items():
        val = data.get(field)
        if val is not None:
            rows.append({"項目": label_names.get(field, field), "値": val, "行": row})

    # Activity Rings extras
    if workout_type and "activity" in workout_type.lower():
        for field in ["move_goal_kcal", "exercise_goal_min", "stand_goal_hours", "step_count", "step_distance_km"]:
            val = data.get(field)
            if val is not None and not any(r["項目"] == label_names.get(field) for r in rows):
                rows.append({"項目": label_names.get(field, field), "値": val, "行": "-"})

    return rows
