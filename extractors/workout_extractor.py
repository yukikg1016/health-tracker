"""
Workout extractor — Apple Fitness app screenshot
Supports: General workout, Outdoor Run, Outdoor Walk, Activity Rings
"""
from extractors.base import call_ai

PROMPT = """
You are analyzing a screenshot from the Apple Fitness app showing workout or activity data.

Extract ALL available fields and return a single JSON object with these keys (use null if not visible):

For workout summaries (General / Outdoor Run / Outdoor Walk):
- workout_type: string (e.g. "Outdoor Run", "Outdoor Walk", "HIIT", "Strength Training", etc.)
- workout_time_min: number (workout duration in minutes)
- elapsed_time_min: number (elapsed/total time in minutes, if different from workout_time)
- distance_km: number (distance in km, convert miles to km if needed)
- active_kcal: number (active calories burned)
- total_kcal: number (total calories burned)
- avg_heart_rate_bpm: number
- avg_pace_min_per_km: number (pace in min/km, convert min/mile if needed)
- avg_power_watts: number
- avg_cadence_spm: number (steps per minute)
- effort: number (effort score 1-10 or as shown)

For Activity Rings:
- move_kcal: number (Move ring — active calories, goal)
- move_goal_kcal: number
- exercise_min: number (Exercise ring minutes)
- exercise_goal_min: number
- stand_hours: number (Stand ring hours)
- stand_goal_hours: number
- step_count: number
- step_distance_km: number (convert miles to km if needed)

Return ONLY valid JSON, no explanation.
""".strip()


def extract_workout_data(image_path: str, provider: str = "gemini", api_key: str = "") -> dict:
    return call_ai(PROMPT, image_path, provider, api_key)
