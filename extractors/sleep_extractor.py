from __future__ import annotations

"""Sleep / HeartWatch スクリーンショット抽出（autosleep_extractor.py を再利用）"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
from autosleep_extractor import extract_sleep_data  # noqa: E402

__all__ = ["extract_sleep_data"]
