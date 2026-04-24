from __future__ import annotations

"""共通AI呼び出しユーティリティ"""

import base64
import json
import os
import sys
from pathlib import Path

# autosleep_extractor.py の関数を再利用
sys.path.insert(0, str(Path(__file__).parent.parent))
from autosleep_extractor import _encode_image, _parse_response  # noqa: E402


def call_ai(prompt: str, image_path: str | None, provider: str, api_key: str) -> dict:
    """
    画像とプロンプトをAI APIに送り、JSONをパースして返す。
    image_path=None の場合はテキストのみモード。

    Args:
        prompt:     抽出プロンプト
        image_path: 画像ファイルのパス（Noneの場合はテキストのみ）
        provider:   "gemini" または "claude"
        api_key:    APIキー

    Returns:
        パース済みdict
    """
    has_image = image_path is not None
    if has_image:
        image_b64, mime_type = _encode_image(image_path)
    else:
        image_b64, mime_type = None, None

    if provider == "gemini":
        try:
            from google import genai
            from google.genai import types
        except ImportError:
            raise ImportError("pip install google-genai")
        client = genai.Client(api_key=api_key)
        if has_image:
            contents = [
                types.Part.from_bytes(data=base64.b64decode(image_b64), mime_type=mime_type),
                prompt,
            ]
        else:
            contents = [prompt]
        # 実際に使えるモデルをAPIから取得して選択
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
            raise RuntimeError("利用可能なGeminiモデルが見つかりません。ListModels結果が空です。")
        model_name = available[0]
        response = client.models.generate_content(model=model_name, contents=contents)
        return _parse_response(response.text)

    elif provider == "claude":
        try:
            import anthropic
        except ImportError:
            raise ImportError("pip install anthropic")
        client = anthropic.Anthropic(api_key=api_key)
        if has_image:
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
            max_tokens=2048,
            messages=[{"role": "user", "content": content}],
        )
        return _parse_response(response.content[0].text)

    else:
        raise ValueError(f"未対応のプロバイダー: {provider}")
