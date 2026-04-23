from __future__ import annotations

"""InBody 体組成スクリーンショット抽出"""

from .base import call_ai

INBODY_PROMPT = """
This is an InBody body composition analysis result screenshot (or similar body composition device).
Extract all visible measurements.

Return ONLY a JSON object in this exact format. Use null for missing values.
Numbers only — no units in values.

{
  "date": "<YYYY-MM-DD or null>",
  "weight":             <float or null>,
  "bmi":                <float or null>,
  "body_fat_pct":       <float or null>,
  "muscle_mass":        <float or null>,
  "body_fat_mass":      <float or null>,
  "visceral_fat_level": <integer or null>,
  "total_body_water":   <float or null>,
  "bone_mass":          <float or null>,
  "basal_metabolic_rate": <integer or null>,
  "skeletal_muscle_mass": <float or null>
}

Return JSON only. No explanation, no markdown fences.
"""


def extract_inbody_data(image_path: str, provider: str, api_key: str) -> dict:
    """
    InBodyスクリーンショットから体組成データを抽出する。
    """
    return call_ai(INBODY_PROMPT, image_path, provider, api_key)
