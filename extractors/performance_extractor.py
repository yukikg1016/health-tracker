"""
Performance extractor — Push up / Chin up / Squat counts from text or screenshot
"""
from extractors.base import call_ai

PROMPT = """
You are analyzing a fitness performance test record.
The user may provide a screenshot, photo, or text with Push up / Chin up / Squat counts.

Extract the following fields:

- push_up: number of push-ups (腕立て伏せ / Push up count)
- chin_up: number of chin-ups (懸垂 / Chin up count)
- squat: number of squats (スクワット / Squat count)

If a field is not visible or mentioned, return null for that field.

Return ONLY this JSON structure:
{
  "workout_type": "Performance",
  "push_up": <integer or null>,
  "chin_up": <integer or null>,
  "squat": <integer or null>
}

Return JSON only. No explanation, no markdown fences.
""".strip()


def extract_performance_data(image_path: str | None, provider: str = "gemini",
                              api_key: str = "", text: str = "") -> dict:
    """
    Extract performance data from an image and/or text input.
    If image_path is None, passes text as the image description.
    """
    if image_path is None and text:
        # Text-only mode: inject the text into the prompt
        combined_prompt = PROMPT + f"\n\nUser input:\n{text}"
        # Call AI with a placeholder — base.call_ai handles None image
        return call_ai(combined_prompt, None, provider, api_key)
    return call_ai(PROMPT, image_path, provider, api_key)
