from __future__ import annotations

"""NutritionシートへのExcel書き込み（行=項目、列=日付 の形式）"""
# v2 — meal_time support added

import datetime
from .writer_base import ExcelWriter

SHEET_NAME = "Nutrition"

# メイン表示行（Column A のラベルに対応）
MEAL_ROW = {
    "breakfast": 2,
    "lunch":     3,
    "dinner":    4,
    "snacks":    5,
}

NUTRITION_ROW_MAP = {
    "total_calories":   6,
    "protein":          11,
    "carbohydrates":    13,
    "fat":              15,
}

MEAL_LABELS = {
    "breakfast": "Breakfast",
    "lunch":     "Lunch",
    "dinner":    "Dinner",
    "snacks":    "Snacks",
}

# 食事別マクロ追跡行（シート下部の非表示領域 row 34+）
# 倍算を防ぐため、各食事の値を個別に記録→合計する
_TRACKING_BASE = 34  # row 34 から開始
_TRACKING_ROWS = {
    ("breakfast", "protein"):      34,
    ("breakfast", "carbohydrates"): 35,
    ("breakfast", "fat"):          36,
    ("lunch",     "protein"):      37,
    ("lunch",     "carbohydrates"): 38,
    ("lunch",     "fat"):          39,
    ("dinner",    "protein"):      40,
    ("dinner",    "carbohydrates"): 41,
    ("dinner",    "fat"):          42,
    ("snacks",    "protein"):      43,
    ("snacks",    "carbohydrates"): 44,
    ("snacks",    "fat"):          45,
}


def find_time_row(ws, meal_type: str) -> int | None:
    """列Aを走査して 'Breakfast time' / 'Lunch time' / 'Dinner time' 等の行番号を返す。"""
    meal_key = meal_type.lower()
    search_terms = {
        "breakfast": ["breakfast time", "朝食 time", "朝time", "朝食時間"],
        "lunch":     ["lunch time", "昼食 time", "昼time", "昼食時間"],
        "dinner":    ["dinner time", "夕食 time", "夜time", "夕食時間"],
        "snacks":    ["snack time", "間食 time", "snacks time", "間食時間"],
    }.get(meal_key, [])

    for row in range(1, ws.max_row + 1):
        val = ws.cell(row, 1).value
        if val is None:
            continue
        val_l = str(val).lower().strip()
        if any(t in val_l for t in search_terms):
            return row
    return None


def _ensure_tracking_labels(ws) -> None:
    """追跡行のラベルが存在しなければ初期化する（A列）。"""
    labels = {
        34: "_breakfast_protein",
        35: "_breakfast_carbs",
        36: "_breakfast_fat",
        37: "_lunch_protein",
        38: "_lunch_carbs",
        39: "_lunch_fat",
        40: "_dinner_protein",
        41: "_dinner_carbs",
        42: "_dinner_fat",
        43: "_snacks_protein",
        44: "_snacks_carbs",
        45: "_snacks_fat",
    }
    for row, label in labels.items():
        if ws.cell(row, 1).value is None:
            ws.cell(row, 1).value = label


def _recalc_day_macros(ws, col: int) -> None:
    """追跡行から1日のマクロ合計を再計算して rows 11/13/15 に書き込む。"""
    protein = sum(
        (ws.cell(_TRACKING_ROWS[(meal, "protein")], col).value or 0)
        for meal in MEAL_ROW
    )
    carbs = sum(
        (ws.cell(_TRACKING_ROWS[(meal, "carbohydrates")], col).value or 0)
        for meal in MEAL_ROW
    )
    fat = sum(
        (ws.cell(_TRACKING_ROWS[(meal, "fat")], col).value or 0)
        for meal in MEAL_ROW
    )
    ws.cell(NUTRITION_ROW_MAP["protein"],       col).value = round(protein, 1) if protein else None
    ws.cell(NUTRITION_ROW_MAP["carbohydrates"], col).value = round(carbs, 1)   if carbs   else None
    ws.cell(NUTRITION_ROW_MAP["fat"],           col).value = round(fat, 1)     if fat     else None


SUPPLEMENT_NAMES = [
    "Mega Creatine",
    "Super Fish Oil",
    "Vitamin C",
    "Vitamin B",
    "Zinc",
    "Vitamin D",
    "Ashwagandha",
    "Magnesium +Chelate",
    "Melatonin",
    "Coenzme Q10",
    "D3 + K2 (4000NE)",
]


def write_supplement_data(taken: list, target_date, excel_path: str) -> int:
    """チェックされたサプリを Excel に書き込む。戻り値: 書き込んだ件数"""
    import datetime as _dt
    writer = ExcelWriter(excel_path)
    wb = writer.load()
    ws = wb[SHEET_NAME]

    if isinstance(target_date, str):
        target_date = _dt.date.fromisoformat(target_date)

    col_idx = ExcelWriter.find_date_column(ws, target_date)
    if col_idx is None:
        col_idx = ws.max_column + 1
        ws.cell(row=1, column=col_idx).value = target_date.strftime("%Y.%m.%d")

    count = 0
    for row in range(1, ws.max_row + 1):
        label = ws.cell(row, 1).value
        if label in taken:
            ws.cell(row, col_idx).value = "✓"
            count += 1

    writer.save(wb)
    return count


def build_nutrition_preview(data: dict, meal_type: str, col_idx: int) -> list[dict]:
    meal_key = meal_type.lower()
    meal_row = MEAL_ROW.get(meal_key, 2)
    calories = data.get("calories_kcal")

    return [
        {"項目": f"{MEAL_LABELS.get(meal_key, meal_type)} カロリー",
         "推定値": calories,
         "書き込み先": f"行{meal_row} 列{col_idx}"},
        {"項目": "Protein (g)",
         "推定値": data.get("protein_g"),
         "書き込み先": f"行{NUTRITION_ROW_MAP['protein']} 列{col_idx}（再計算）"},
        {"項目": "Carbohydrates (g)",
         "推定値": data.get("carbohydrates_g"),
         "書き込み先": f"行{NUTRITION_ROW_MAP['carbohydrates']} 列{col_idx}（再計算）"},
        {"項目": "Fat (g)",
         "推定値": data.get("fat_g"),
         "書き込み先": f"行{NUTRITION_ROW_MAP['fat']} 列{col_idx}（再計算）"},
    ]


def write_nutrition_data(
    data: dict,
    target_date: datetime.date,
    meal_type: str,
    excel_path: str,
    meal_time: str = "",
) -> list[dict]:
    """
    Nutritionシートに書き込む。

    - 食事カロリーを meal 行（Breakfast=2 etc.）に上書き
    - per-meal マクロを追跡行（row34-45）に上書き（倍算防止）
    - 1日のマクロ合計を行11/13/15 に再計算
    - Total Calories（行6）= 行2-5 の合計

    Returns:
        プレビュー行リスト
    """
    writer = ExcelWriter(excel_path)
    wb = writer.load()
    ws = wb[SHEET_NAME]

    # 日付列を探す（なければ末尾に追加）
    col_idx = ExcelWriter.find_date_column(ws, target_date)
    if col_idx is None:
        col_idx = ws.max_column + 1
        ws.cell(row=1, column=col_idx).value = target_date.strftime("%Y.%m.%d")

    preview = build_nutrition_preview(data, meal_type, col_idx)

    meal_key = meal_type.lower()
    meal_row = MEAL_ROW.get(meal_key, 2)

    # ① カロリーを食事行に上書き
    calories = data.get("calories_kcal")
    ws.cell(row=meal_row, column=col_idx).value = round(calories) if calories is not None else None

    # ② Total Calories を再計算（行2-5の合計）
    total = sum(
        (ws.cell(r, col_idx).value or 0)
        for r in range(2, 6)
        if isinstance(ws.cell(r, col_idx).value, (int, float))
    )
    ws.cell(row=NUTRITION_ROW_MAP["total_calories"], column=col_idx).value = round(total) if total else None

    # ③ 追跡行のラベル初期化
    _ensure_tracking_labels(ws)

    # ④ この食事のマクロを追跡行に上書き（倍算防止）
    for macro, key in [("protein", "protein_g"), ("carbohydrates", "carbohydrates_g"), ("fat", "fat_g")]:
        track_row = _TRACKING_ROWS[(meal_key, macro)]
        val = data.get(key)
        ws.cell(row=track_row, column=col_idx).value = round(val, 1) if val is not None else None

    # ⑤ 1日のマクロ合計を再計算して行11/13/15に書き込み
    _recalc_day_macros(ws, col_idx)

    # ⑥ 食事時間を書き込む（列Aからラベル行を検索）
    if meal_time and meal_time.strip():
        time_row = find_time_row(ws, meal_type)
        if time_row is not None:
            ws.cell(row=time_row, column=col_idx).value = meal_time.strip()

    writer.save(wb)
    return preview
