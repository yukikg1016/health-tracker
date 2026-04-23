from __future__ import annotations

"""夜💤シート用スクリーンショット抽出"""

from .base import call_ai

NIGHT_PROMPT = """
This is an evening health/wellness log screenshot. Extract the visible data.

Return ONLY a JSON object in this exact format. Use null for missing values.

{
  "date": "<YYYY-MM-DD or null>",
  "water_intake_ml":  <integer or null>,
  "training_rpe":     <float or null>,
  "study_hours":      <float or null>,
  "jump_cm":          <float or null>,
  "coffee_cups":      <float or null>,
  "run_distance_km":  <float or null>,
  "run_power_w":      <float or null>,
  "pushups":          <integer or null>,
  "grip_strength_kg": <float or null>,
  "notes":            "<text or null>"
}

Field descriptions:
- water_intake_ml: total daily water intake in milliliters
- training_rpe: training intensity on RPE scale 1-10
- study_hours: study / deep work hours
- jump_cm: vertical jump record in centimeters
- coffee_cups: coffee cups consumed (mark ⭐ if after 4pm, but just extract number)
- run_distance_km: 10-minute run distance in km
- run_power_w: 10-minute run power in watts
- pushups: pushup count
- grip_strength_kg: grip strength in kg
- notes: any additional notes

Return JSON only. No explanation, no markdown fences.
"""


def extract_night_data(image_path: str, provider: str, api_key: str) -> dict:
    """
    夜の記録スクリーンショットからデータを抽出する。
    """
    return call_ai(NIGHT_PROMPT, image_path, provider, api_key)
