from __future__ import annotations

"""朝☀️シート用スクリーンショット抽出"""

from .base import call_ai

MORNING_PROMPT = """
This is a morning health/wellness survey or app screenshot. Extract the visible data.

Return ONLY a JSON object in this exact format. Use null for missing values.

{
  "date": "<YYYY-MM-DD or null>",
  "reaction_time_sec":   <float or null>,
  "alcohol_drinks":      <integer or null>,
  "coffee_after4pm":     <float or null>,
  "fatigue_1to10":       <integer or null>,
  "stress_1to10":        <integer or null>,
  "device_before_bed":   <"yes" | "no" | null>,
  "mood_1to10":          <integer or null>,
  "accuracy_pct":        <float or null>,
  "sleep_duration_text": "<e.g. 7時間29分 or null>",
  "deep_sleep_text":     "<e.g. 1時間45分 or null>",
  "sleep_score":         <float or null>,
  "test_result_s":       <float or null>
}

Field descriptions:
- reaction_time_sec: Stop Signal Test reaction time in seconds
- alcohol_drinks: number of alcoholic drinks last night
- coffee_after4pm: cups of coffee after 4pm yesterday
- fatigue_1to10: subjective fatigue level 1-10
- stress_1to10: subjective stress level 1-10
- device_before_bed: did they use devices 30 minutes before bed
- mood_1to10: overall mood rating
- accuracy_pct: Stop Signal Test accuracy percentage
- sleep_duration_text: sleep duration as text (keep as-is)
- deep_sleep_text: deep sleep duration as text (keep as-is)
- sleep_score: sleep quality score
- test_result_s: Duali Real World test result in seconds

Return JSON only. No explanation, no markdown fences.
"""


def extract_morning_data(image_path: str, provider: str, api_key: str) -> dict:
    """
    朝の記録スクリーンショットからデータを抽出する。
    """
    return call_ai(MORNING_PROMPT, image_path, provider, api_key)
