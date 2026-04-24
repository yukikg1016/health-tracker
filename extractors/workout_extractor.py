"""
Workout extractor — Apple Fitness app screenshot
Supports: General workout, Outdoor Run, Outdoor Walk, Activity Rings
"""
from extractors.base import call_ai

PROMPT = """
You are analyzing a screenshot from the Apple Fitness app (Japanese UI).

Extract ALL visible fields and return a single JSON object. Use null for any field not visible.

--- WORKOUT SCREENSHOT (individual workout detail page) ---
Japanese label → JSON key:
- ワークアウト時間 → workout_time_min  (format H:MM:SS → convert to decimal minutes, e.g. 0:15:45 → 15.75)
- 経過時間 → elapsed_time_min  (same conversion)
- 距離 → distance_km  (convert M→km by /1000, e.g. 533M→0.533, 1.24KM→1.24)
- アクティブキロカロリー → active_kcal  (number only)
- 合計キロカロリー → total_kcal  (number only)
- 平均心拍数 → avg_heart_rate_bpm  (number before "拍/分")
- 平均ペース → avg_pace_min_per_km  (format MM'SS"/KM → decimal min/km, e.g. 12'38" → 12.633)
- 平均パワー → avg_power_watts
- 平均ケイデンス → avg_cadence_spm
- 難易度 → effort  (extract the NUMBER only, e.g. "2 簡単" → 2, "8 きつい" → 8)

Workout type (top of screen, in Japanese):
- ウォーキング（屋外） → "Outdoor Walk"
- ランニング（屋外） → "Outdoor Run"
- コアトレーニング → "Core Training"
- ストレングストレーニング → "Strength Training"
- その他 → use the Japanese name as-is

Set these Activity Ring fields to null for individual workout screenshots:
move_kcal, move_goal_kcal, exercise_min, exercise_goal_min, stand_hours, stand_goal_hours, step_count, step_distance_km

--- ACTIVITY RINGS SCREENSHOT (summary/概要 page) ---
Japanese label → JSON key:
- ムーブ  "XX/YY KCAL" → move_kcal=XX (current), move_goal_kcal=YY
- エクササイズ "XX/YY 分" → exercise_min=XX, exercise_goal_min=YY
- ロール or スタンド "XX/YY 時間" → stand_hours=XX, stand_goal_hours=YY
- 歩数 → step_count
- 歩行距離 → step_distance_km (convert M→km if needed)

Set workout_type="Activity Rings" and all workout fields (workout_time_min etc.) to null for Activity Rings screenshots.

Return ONLY this JSON structure:
{
  "workout_type": "<string or null>",
  "workout_time_min": <float or null>,
  "elapsed_time_min": <float or null>,
  "distance_km": <float or null>,
  "active_kcal": <float or null>,
  "total_kcal": <float or null>,
  "avg_heart_rate_bpm": <float or null>,
  "avg_pace_min_per_km": <float or null>,
  "avg_power_watts": <float or null>,
  "avg_cadence_spm": <float or null>,
  "effort": <integer or null>,
  "move_kcal": <float or null>,
  "move_goal_kcal": <float or null>,
  "exercise_min": <float or null>,
  "exercise_goal_min": <float or null>,
  "stand_hours": <float or null>,
  "stand_goal_hours": <float or null>,
  "step_count": <integer or null>,
  "step_distance_km": <float or null>
}

Return JSON only. No explanation, no markdown fences.
""".strip()


def extract_workout_data(image_path: str, provider: str = "gemini", api_key: str = "") -> dict:
    return call_ai(PROMPT, image_path, provider, api_key)
