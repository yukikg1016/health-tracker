from __future__ import annotations

"""血液検査（Labs）スクリーンショット抽出"""

from .base import call_ai

LABS_PROMPT = """
This is a blood test / lab report screenshot. Extract all visible test results.

Return ONLY a JSON object in this exact format. Use null for missing values.

{
  "test_date": "<YYYY-MM-DD or null>",
  "results": [
    {"name": "<test name exactly as shown>", "value": <number or null>, "unit": "<unit or null>"},
    ...
  ]
}

Common test names to look for (use these exact spellings if they appear):
Sodium, Potassium, Chloride, Bicarbonate, Anion Gap, Calcium, Corrected Calcium, Magnesium,
Glucose, Haemoglobin A1C, Total Protein, Albumin, Cholesterol, Triglycerides,
HDL-Cholesterol, LDL-Cholesterol, Sensitive C-Reactive Protein, Erythrocyte Sedimentation Rate,
AST/GOT, ALT/GPT, ALP, GGT, Total Bilirubin, Alpha-Amylase, Lipase,
Urea, Creatinine, eGFR-EPI, Iron, Ferritin,
WBC Count, Neutrophil, Lymphocyte, Monocyte, Eosinophil, Basophil,
RBC Count, Haemoglobin, Hematocrit, MCV, MCH, MCHC, RDW-CV,
Platelet Count, MPV, Folate, Vitamin B12, TSH, Testosterone, Estradiol, Cortisol

Return JSON only. No explanation, no markdown fences.
"""


def extract_labs_data(image_path: str, provider: str, api_key: str) -> dict:
    """
    血液検査スクリーンショットから検査値を抽出する。

    Returns:
        {"test_date": "YYYY-MM-DD or null", "results": [{"name": ..., "value": ..., "unit": ...}]}
    """
    return call_ai(LABS_PROMPT, image_path, provider, api_key)
