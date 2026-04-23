from __future__ import annotations

"""Nutrition スクリーンショット + テキスト説明からの栄養値抽出"""

import base64
import json
import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from autosleep_extractor import _encode_image, _parse_response


NUTRITION_PROMPT_TEMPLATE = """
You are a professional nutritionist AI. Analyze the food {input_type} below to estimate all nutritional values as accurately as possible.

User description: {description}

Return ONLY a JSON object in this exact format. Use null if truly impossible to estimate — but always try your best to give a numeric estimate.
Do NOT include units in values — numbers only.

{{
  "meal": "<breakfast | lunch | dinner | snack>",
  "description_summary": "<1-line summary of the food>",
  "calories_kcal":      <float>,
  "protein_g":          <float>,
  "fat_g":              <float>,
  "carbohydrates_g":    <float>,
  "sugar_g":            <float>,
  "dietary_fiber_g":    <float>,
  "saturated_fat_g":    <float>,
  "sodium_mg":          <float>,
  "potassium_mg":       <float>,
  "calcium_mg":         <float>,
  "iron_mg":            <float>,
  "vitamin_c_mg":       <float>,
  "notes": "<estimation notes, e.g. portion size assumptions or null>"
}}

Important:
- Base estimates on typical Japanese/Western food databases
- If portion size is unclear, assume a standard single serving
- If the image shows a restaurant meal, account for typical restaurant portions
- Return JSON only. No explanation, no markdown fences.
"""


def extract_nutrition_data(
    image_path: str | None,
    provider: str,
    api_key: str,
    description: str = "",
) -> dict:
    """
    食事の画像（省略可）＋テキスト説明から栄養値を推定する。

    Args:
        image_path:   食事画像のパス（None の場合はテキストのみで推定）
        provider:     "gemini" または "claude"
        api_key:      APIキー
        description:  ユーザーの補足説明（食材、量、料理名など）

    Returns:
        推定栄養値の dict
    """
    has_image = image_path is not None
    input_type = "and description" if has_image else "description (no image provided)"
    prompt = NUTRITION_PROMPT_TEMPLATE.format(
        description=description if description.strip() else "（説明なし）",
        input_type=input_type,
    )

    if provider == "gemini":
        try:
            from google import genai
            from google.genai import types
        except ImportError:
            raise ImportError("pip install google-genai")
        client = genai.Client(api_key=api_key)
        if has_image:
            image_b64, mime_type = _encode_image(image_path)
            contents = [
                types.Part.from_bytes(data=base64.b64decode(image_b64), mime_type=mime_type),
                prompt,
            ]
        else:
            contents = [prompt]
        available = [
            m.name for m in client.models.list()
            if "generateContent" in (m.supported_actions or [])
            and "gemini" in m.name.lower()
            and "flash" in m.name.lower()
        ]
        if not available:
            available = [
                m.name for m in client.models.list()
                if "generateContent" in (m.supported_actions or [])
                and "gemini" in m.name.lower()
            ]
        if not available:
            raise RuntimeError("利用可能なGeminiモデルが見つかりません。")
        response = client.models.generate_content(model=available[0], contents=contents)
        return _parse_response(response.text)

    elif provider == "claude":
        try:
            import anthropic
        except ImportError:
            raise ImportError("pip install anthropic")
        client = anthropic.Anthropic(api_key=api_key)
        if has_image:
            image_b64, mime_type = _encode_image(image_path)
            content = [
                {"type": "image", "source": {
                    "type": "base64", "media_type": mime_type, "data": image_b64,
                }},
                {"type": "text", "text": prompt},
            ]
        else:
            content = [{"type": "text", "text": prompt}]
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1024,
            messages=[{"role": "user", "content": content}],
        )
        return _parse_response(response.content[0].text)

    else:
        raise ValueError(f"未対応のプロバイダー: {provider}")
