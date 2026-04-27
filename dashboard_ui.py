"""
Dashboard UI — Recovery Score, Training Recommendation, Biological Age
"""
from __future__ import annotations
import datetime
import json


# ── Recovery Score計算 ────────────────────────────────────────────────────────

def calc_recovery_score(sleep_records: list[dict]) -> dict:
    """
    Calculate today's recovery score (0-100) from recent sleep data.
    Returns score, grade, and component breakdown.
    """
    if not sleep_records:
        return {"score": None, "grade": "データなし", "components": {}}

    # Use most recent record
    latest = sleep_records[0]

    components = {}
    total_weight = 0
    total_score = 0

    # 1. Total sleep (weight: 35)
    ts = latest.get("total_sleep_h")
    if ts is not None:
        # Optimal: 7-9h. Score drops outside this range
        if 7 <= ts <= 9:
            s = 100
        elif 6 <= ts < 7 or 9 < ts <= 10:
            s = 75
        elif 5 <= ts < 6 or 10 < ts <= 11:
            s = 45
        else:
            s = 20
        components["睡眠時間"] = {"score": s, "value": f"{ts:.1f}h", "weight": 35}
        total_score += s * 35
        total_weight += 35

    # 2. Deep sleep (weight: 25)
    ds = latest.get("deep_sleep_h")
    if ds is not None:
        if ds >= 1.5:
            s = 100
        elif ds >= 1.0:
            s = 75
        elif ds >= 0.5:
            s = 45
        else:
            s = 20
        components["深い睡眠"] = {"score": s, "value": f"{ds:.1f}h", "weight": 25}
        total_score += s * 25
        total_weight += 25

    # 3. Sleep quality score (weight: 20)
    qs = latest.get("quality_score")
    if qs is not None:
        # Assume score is 0-100 or percentage
        if qs > 10:  # already 0-100
            s = min(100, float(qs))
        else:  # assume 0-10 scale
            s = min(100, float(qs) * 10)
        components["睡眠品質"] = {"score": s, "value": f"{qs:.0f}", "weight": 20}
        total_score += s * 20
        total_weight += 20

    # 4. Resting heart rate (weight: 20)
    rhr = latest.get("resting_hr")
    if rhr is not None:
        if rhr <= 50:
            s = 100
        elif rhr <= 55:
            s = 90
        elif rhr <= 60:
            s = 75
        elif rhr <= 65:
            s = 55
        elif rhr <= 70:
            s = 35
        else:
            s = 20
        components["安静時心拍"] = {"score": s, "value": f"{rhr:.0f} bpm", "weight": 20}
        total_score += s * 20
        total_weight += 20

    if total_weight == 0:
        return {"score": None, "grade": "データなし", "components": {}}

    score = round(total_score / total_weight)

    if score >= 80:
        grade = "excellent"
    elif score >= 65:
        grade = "good"
    elif score >= 45:
        grade = "moderate"
    else:
        grade = "poor"

    return {"score": score, "grade": grade, "components": components,
            "date": latest.get("date")}


def calc_training_recommendation(recovery: dict, workout_records: list[dict]) -> dict:
    """
    Recommend today's training intensity based on recovery score and recent load.
    """
    score = recovery.get("score")

    # Calculate recent training load (last 3 days)
    recent_load = 0
    if workout_records:
        for i, w in enumerate(workout_records[:3]):
            effort = w.get("effort", 3)
            time_min = w.get("workout_time_min", 30)
            weight = 1.0 if i == 0 else (0.7 if i == 1 else 0.4)
            recent_load += (effort / 5) * (time_min / 60) * weight

    if score is None:
        return {
            "intensity": "unknown",
            "label": "データ不足",
            "color": "gray",
            "advice": "睡眠データを記録するとアドバイスが表示されます",
        }

    # Decision matrix
    if score >= 80:
        if recent_load > 1.5:
            return {
                "intensity": "moderate",
                "label": "💪 中〜高強度",
                "color": "orange",
                "advice": "回復は良好ですが、直近の負荷が高め。70〜80%強度で行いましょう",
            }
        return {
            "intensity": "high",
            "label": "🔥 高強度OK",
            "color": "green",
            "advice": "回復スコア優秀。高強度トレーニングに最適な日です",
        }
    elif score >= 65:
        return {
            "intensity": "moderate",
            "label": "💪 中強度推奨",
            "color": "orange",
            "advice": "良好な回復状態。70〜80%強度のトレーニングを推奨",
        }
    elif score >= 45:
        return {
            "intensity": "light",
            "label": "🚶 軽め推奨",
            "color": "yellow",
            "advice": "回復が不十分。ウォーキングやストレッチ、軽いワークアウトに留めましょう",
        }
    else:
        return {
            "intensity": "rest",
            "label": "😴 休養日推奨",
            "color": "red",
            "advice": "回復スコアが低い。今日は積極的休養（軽いストレッチのみ）を推奨します",
        }


def calc_biological_age(inbody: dict | None, sleep_records: list[dict],
                         labs: dict | None, workout_records: list[dict],
                         actual_age: int | None = None) -> dict:
    """
    Estimate biological age from available biomarkers.
    Returns estimated_age, components, and confidence.
    """
    scores = []  # list of (estimated_age_offset, weight, label, detail)

    # 1. Resting HR → VO2max proxy
    if sleep_records:
        rhr = next((r.get("resting_hr") for r in sleep_records if r.get("resting_hr")), None)
        if rhr is not None:
            # Lower RHR = younger biological age
            # Average 30yo male: ~62bpm, 40yo: ~66bpm
            offset = (rhr - 55) * 0.4  # each bpm above 55 ≈ +0.4 years
            scores.append((offset, 3, "安静時心拍", f"{rhr:.0f} bpm"))

    # 2. Body composition
    if inbody:
        bf = inbody.get("body_fat_pct")
        muscle = inbody.get("muscle_kg")
        if bf is not None:
            # Optimal male body fat: 10-15% for athletes
            if bf <= 12:
                offset = -3
            elif bf <= 17:
                offset = 0
            elif bf <= 22:
                offset = 3
            else:
                offset = 7
            scores.append((offset, 3, "体脂肪率", f"{bf:.1f}%"))

        if muscle is not None:
            # High muscle mass → younger
            if muscle >= 35:
                offset = -2
            elif muscle >= 30:
                offset = 0
            else:
                offset = 3
            scores.append((offset, 2, "筋肉量", f"{muscle:.1f} kg"))

    # 3. Sleep quality (avg of recent 7 days)
    if sleep_records:
        recent_sleep = [r.get("total_sleep_h") for r in sleep_records[:7] if r.get("total_sleep_h")]
        if recent_sleep:
            avg = sum(recent_sleep) / len(recent_sleep)
            if avg >= 7.5:
                offset = -2
            elif avg >= 6.5:
                offset = 0
            elif avg >= 5.5:
                offset = 3
            else:
                offset = 6
            scores.append((offset, 2, "平均睡眠時間", f"{avg:.1f}h"))

    # 4. Labs
    if labs:
        vd = labs.get("vitamin_d")
        if vd is not None:
            if vd >= 50:
                offset = -1
            elif vd >= 30:
                offset = 0
            else:
                offset = 2
            scores.append((offset, 1, "ビタミンD", f"{vd:.0f} ng/mL"))

        ferritin = labs.get("ferritin")
        if ferritin is not None:
            if 50 <= ferritin <= 200:
                offset = -1
            elif ferritin < 30:
                offset = 2
            else:
                offset = 0
            scores.append((offset, 1, "フェリチン", f"{ferritin:.0f}"))

    if not scores:
        return {"estimated_age": None, "confidence": 0, "components": []}

    total_weight = sum(w for _, w, _, _ in scores)
    weighted_offset = sum(offset * w for offset, w, _, _ in scores) / total_weight

    # Base: use actual age if provided, otherwise assume 30
    base = actual_age if actual_age else 30
    estimated = round(base + weighted_offset)

    confidence = min(100, int(total_weight / 12 * 100))

    components = [
        {"label": label, "detail": detail,
         "impact": "若い" if offset < 0 else ("標準" if offset == 0 else "老化")}
        for offset, _, label, detail in scores
    ]

    return {
        "estimated_age": estimated,
        "offset": round(weighted_offset, 1),
        "confidence": confidence,
        "components": components,
    }


def build_ai_prompt(data: dict, actual_age: int | None = None) -> str:
    """Build a prompt for AI daily advice generation."""
    sleep = data.get("sleep", [])
    workout = data.get("workout", [])
    nutrition = data.get("nutrition", [])
    inbody = data.get("inbody")
    labs = data.get("labs")

    latest_sleep = sleep[0] if sleep else {}
    latest_nutrition = nutrition[0] if nutrition else {}

    avg_sleep = None
    if sleep:
        vals = [r.get("total_sleep_h") for r in sleep[:7] if r.get("total_sleep_h")]
        if vals:
            avg_sleep = round(sum(vals) / len(vals), 1)

    prompt = f"""
You are an elite athletic performance coach and sports nutritionist.
Analyze the following athlete data and provide personalized daily advice in Japanese.
Today is {datetime.date.today().strftime('%Y年%m月%d日')}.

## Today's Data
- Total sleep: {latest_sleep.get('total_sleep_h', 'N/A')} hours
- Deep sleep: {latest_sleep.get('deep_sleep_h', 'N/A')} hours
- Resting HR: {latest_sleep.get('resting_hr', 'N/A')} bpm
- Sleep quality score: {latest_sleep.get('quality_score', 'N/A')}
- 7-day avg sleep: {avg_sleep or 'N/A'} hours

## Recent Workout (last 3 sessions)
{json.dumps([{k: v for k, v in w.items() if k != 'date'} for w in workout[:3]], ensure_ascii=False)}

## Today's Nutrition
- Calories: {latest_nutrition.get('calories', 'N/A')} kcal
- Protein: {latest_nutrition.get('protein_g', 'N/A')} g
- Carbs: {latest_nutrition.get('carbs_g', 'N/A')} g
- Fat: {latest_nutrition.get('fat_g', 'N/A')} g

## Body Composition (latest)
{json.dumps({k: v for k, v in (inbody or {}).items() if k != 'date'}, ensure_ascii=False)}

## Blood Labs (latest)
{json.dumps({k: v for k, v in (labs or {}).items() if k != 'date'}, ensure_ascii=False)}

Provide advice in this exact JSON format (Japanese text):
{{
  "today_priority": "<1-sentence most important thing for today>",
  "nutrition_advice": "<specific nutrition tip based on data>",
  "sleep_target": "<tonight's sleep recommendation with specific bedtime>",
  "supplement_timing": "<supplement timing advice for today>",
  "training_note": "<1 specific training tip for today>"
}}

Return JSON only. No markdown fences.
""".strip()
    return prompt


def get_ai_advice(prompt: str, provider: str, api_key: str) -> dict:
    """Call AI to generate daily advice."""
    import sys
    import pathlib
    sys.path.insert(0, str(pathlib.Path(__file__).parent))
    from autosleep_extractor import _parse_response

    if provider == "gemini":
        from google import genai
        client = genai.Client(api_key=api_key)
        available = [
            m.name for m in client.models.list()
            if "generateContent" in (m.supported_actions or [])
            and "gemini" in m.name.lower() and "flash" in m.name.lower()
        ]
        if not available:
            available = [m.name for m in client.models.list()
                         if "generateContent" in (m.supported_actions or [])
                         and "gemini" in m.name.lower()]
        response = client.models.generate_content(model=available[0], contents=[prompt])
        return _parse_response(response.text)

    elif provider == "claude":
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        response = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        return _parse_response(response.content[0].text)

    return {}
