"""
AutoSleep / HeartWatch Screenshot Data Extractor

使い方:
  # 1. CLIから実行
  python autosleep_extractor.py screenshot.png
  python autosleep_extractor.py screenshot.png --provider claude --output result.json

  # 2. Pythonからimportして使う
  from autosleep_extractor import extract_sleep_data
  data = extract_sleep_data("screenshot.png", provider="gemini")

  # 3. ファイル上部の変数を書き換えて直接実行
  IMAGE_PATH = "screenshot.png"
  python autosleep_extractor.py

インストール:
  pip install google-generativeai   # Gemini を使う場合
  pip install anthropic              # Claude を使う場合
"""

from __future__ import annotations

import base64
import json
import os
import sys
from pathlib import Path

# ─────────────────────────────────────────────
# ★ CLI引数なしで直接実行する場合はここを編集
IMAGE_PATH = "screenshot.png"   # 画像パス（リストで複数指定も可）
PROVIDER   = "gemini"           # "gemini" または "claude"
API_KEY    = ""                 # 空のとき環境変数 GEMINI_API_KEY / ANTHROPIC_API_KEY を使用
OUTPUT     = "sleep_data.json"  # 出力ファイル（空文字 "" で標準出力のみ）
# ─────────────────────────────────────────────

PROMPT = """
This is a screenshot from the AutoSleep / HeartWatch app. Extract all sleep data visible on screen.

Return ONLY a JSON object in the exact format below. Use null for any field not visible.
- Duration fields: split into hours, minutes, total_minutes (e.g. "7h 23m" → hours:7, minutes:23, total_minutes:443)
- Time-of-day fields (HH:MM): keep as string
- BPM / SpO2 / HRV / percentage / scores: numeric only, no units
- If the note "※Sleep data is from prior evening" is visible, set prior_evening to true

{
  "date": "<YYYY-MM-DD or null>",
  "prior_evening": <true | false>,

  "hearwatch": {
    "daily_bpm":              <integer or null>,
    "sedentary_bpm":          <integer or null>,
    "sleep_bpm":              <integer or null>,
    "sleep_spo2":             <float or null>,
    "sleep_respiratory_rate": <float or null>,
    "sleep_time": {
      "hours":         <integer or null>,
      "minutes":       <integer or null>,
      "total_minutes": <integer or null>
    },
    "restfulness":  <float or null>,
    "sleep_hrv":    <integer or null>,
    "waking_bpm":   <integer or null>,
    "waking_hrv":   <integer or null>,
    "daily_spo2":   <float or null>
  },

  "sleep": {
    "total_sleep": {
      "hours":         <integer or null>,
      "minutes":       <integer or null>,
      "total_minutes": <integer or null>
    },
    "sleep_quality":       <float or null>,
    "deep_sleep": {
      "hours":         <integer or null>,
      "minutes":       <integer or null>,
      "total_minutes": <integer or null>,
      "percentage":    <float or null>
    },
    "bpm_sleep_avg":       <float or null>,
    "sleep_session_start": "<HH:MM or null>",
    "sleep_session_end":   "<HH:MM or null>",
    "sleep_efficiency":    <float or null>,
    "sleep_rating":        <float or null>
  },

  "wellness": {
    "readiness_score":         <float or null>,
    "sleep_fuel_rating":       "<text label e.g. プレミアム/良い/普通/不足 or null>",
    "hrv":                     <float or null>,
    "bpm_hrv":                 <integer or null>,
    "prior_day_stress":        "<text label e.g. 良い/普通/不足 or null>",
    "stress_hrv":              <float or null>,
    "sleep_bank":              "<text or float or null>",
    "sleep_bank_balance":      "<text or float e.g. '34.2% 借金' or null>",
    "wrist_temp":              <float or null>,
    "wrist_temp_baseline":     <float or null>,
    "wrist_temp_deviation":    <float or null>,
    "sleep_spo2_avg":          <float or null>,
    "sleep_spo2_range":        "<text e.g. '95% - 98%' or null>",
    "respiration_rate_avg":    <float or null>,
    "respiration_rate_range":  "<text e.g. '17.0 - 35.5' or null>"
  }
}

Field mapping hints (Japanese → JSON key):
- 睡眠燃料評価 → sleep_fuel_rating
- 今日の快適さ / Readiness → readiness_score (numeric if shown, else text)
- 前日のストレス → prior_day_stress (label text)
- 前日のストレスの心拍変動 → stress_hrv (numeric ms value)
- 借金% / Sleep Bank Balance → sleep_bank_balance
- 温度 (actual value) → wrist_temp
- 温度ベースライン → wrist_temp_baseline
- 温度偏差 (↑↓ value) → wrist_temp_deviation (negative if ↓)
- 睡眠SpO2 平均 → sleep_spo2_avg
- 睡眠SpO2 範囲 → sleep_spo2_range
- 呼吸数 平均 → respiration_rate_avg
- 呼吸数 範囲 → respiration_rate_range

Return JSON only. No explanation, no markdown fences.
"""


def _encode_image(image_path: str) -> tuple[str, str]:
    suffix = Path(image_path).suffix.lower()
    mime_type = {".jpg": "image/jpeg", ".jpeg": "image/jpeg"}.get(suffix, "image/png")
    with open(image_path, "rb") as f:
        return base64.standard_b64encode(f.read()).decode("utf-8"), mime_type


def _parse_response(text: str) -> dict:
    text = text.strip()
    if text.startswith("```"):
        lines = text.splitlines()
        text = "\n".join(lines[1:-1] if lines[-1].strip() == "```" else lines[1:]).strip()
    return json.loads(text)


def extract_sleep_data(
    image_path: str,
    provider: str = "gemini",
    api_key: str | None = None,
) -> dict:
    """
    AutoSleep / HeartWatch スクリーンショットから睡眠データを抽出する

    Args:
        image_path: スクリーンショットのファイルパス
        provider:   "gemini" または "claude"
        api_key:    APIキー（省略時は環境変数 GEMINI_API_KEY / ANTHROPIC_API_KEY を使用）

    Returns:
        睡眠データの dict（hearwatch / sleep の2セクション構成）
    """
    if not Path(image_path).exists():
        raise FileNotFoundError(f"画像が見つかりません: {image_path}")

    key = api_key or os.environ.get(
        "GEMINI_API_KEY" if provider == "gemini" else "ANTHROPIC_API_KEY", ""
    )
    if not key:
        raise ValueError(
            f"APIキーが設定されていません。"
            f"引数 api_key を指定するか、環境変数 "
            f"{'GEMINI_API_KEY' if provider == 'gemini' else 'ANTHROPIC_API_KEY'} を設定してください。"
        )

    image_b64, mime_type = _encode_image(image_path)

    if provider == "gemini":
        try:
            from google import genai
            from google.genai import types
        except ImportError:
            raise ImportError("pip install google-genai")
        client = genai.Client(api_key=key)
        contents = [
            types.Part.from_bytes(data=base64.b64decode(image_b64), mime_type=mime_type),
            PROMPT,
        ]
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
        client = anthropic.Anthropic(api_key=key)
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1024,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {
                        "type": "base64", "media_type": mime_type, "data": image_b64,
                    }},
                    {"type": "text", "text": PROMPT},
                ],
            }],
        )
        return _parse_response(response.content[0].text)

    else:
        raise ValueError(f"未対応のプロバイダー: {provider}（'gemini' または 'claude'）")


def process_multiple(
    image_paths: list[str],
    provider: str = "gemini",
    api_key: str | None = None,
) -> list[dict]:
    """複数のスクリーンショットをまとめて処理する"""
    results = []
    for path in image_paths:
        print(f"処理中: {path} ...", file=sys.stderr)
        try:
            data = extract_sleep_data(path, provider, api_key)
            data["_source_file"] = path
            data["_status"] = "success"
        except Exception as e:
            data = {"_source_file": path, "_status": "error", "_error": str(e)}
            print(f"  → エラー: {e}", file=sys.stderr)
        else:
            print("  → 完了", file=sys.stderr)
        results.append(data)
    return results


# ── CLI / 直接実行 ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="AutoSleep / HeartWatch スクリーンショットから睡眠データをJSON抽出"
    )
    parser.add_argument("images", nargs="*", help="スクリーンショットのパス（省略時はファイル上部の IMAGE_PATH を使用）")
    parser.add_argument("--provider", "-p", choices=["gemini", "claude"], help="使用するAI API")
    parser.add_argument("--api-key",  "-k", help="APIキー（省略時は環境変数を使用）")
    parser.add_argument("--output",   "-o", help="出力JSONファイルのパス")
    parser.add_argument("--indent",         type=int, default=2, help="JSONインデント数（デフォルト: 2）")
    args = parser.parse_args()

    # CLI引数 > ファイル上部の変数 の優先順位で解決
    paths    = args.images   or (IMAGE_PATH if isinstance(IMAGE_PATH, list) else [IMAGE_PATH])
    provider = args.provider or PROVIDER
    api_key  = args.api_key  or API_KEY or None
    output   = args.output   or OUTPUT

    result = (
        extract_sleep_data(paths[0], provider, api_key)
        if len(paths) == 1
        else process_multiple(paths, provider, api_key)
    )

    output_json = json.dumps(result, ensure_ascii=False, indent=args.indent)
    print(output_json)

    if output:
        Path(output).write_text(output_json, encoding="utf-8")
        print(f"\n保存しました: {output}", file=sys.stderr)
