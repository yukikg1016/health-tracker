from __future__ import annotations

"""LabsシートへのExcel書き込み"""

import datetime
import difflib
from .writer_base import ExcelWriter

FUZZY_CUTOFF = 0.75


def _build_test_name_index(ws) -> dict[str, int]:
    """Col A の検査名 → 行番号 の辞書（正規化済みキー）を返す。"""
    index = {}
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val:
            index[str(val).strip().lower()] = row
    return index


def _fuzzy_match(ai_name: str, index: dict[str, int]) -> tuple[str, int, float] | None:
    """
    ai_name を index のキーに対して fuzzy マッチングし、
    (matched_name, row_idx, confidence) を返す。マッチなしは None。
    """
    normalized = ai_name.strip().lower()
    # 完全一致
    if normalized in index:
        return normalized, index[normalized], 1.0
    # fuzzy マッチ
    matches = difflib.get_close_matches(normalized, index.keys(), n=1, cutoff=FUZZY_CUTOFF)
    if matches:
        matched = matches[0]
        ratio = difflib.SequenceMatcher(None, normalized, matched).ratio()
        return matched, index[matched], ratio
    return None


def _find_or_create_date_col(ws, target_date: datetime.date) -> int:
    """
    Row 1 の col B, C を確認し、空なら B から使う。
    すでに日付が入っていれば既存列を返す。新規なら末尾に追加。
    """
    for col in range(2, ws.max_column + 2):
        val = ws.cell(row=1, column=col).value
        if val is None:
            ws.cell(row=1, column=col).value = datetime.datetime(
                target_date.year, target_date.month, target_date.day
            )
            return col
        if isinstance(val, datetime.datetime) and val.date() == target_date:
            return col
        if isinstance(val, datetime.date) and val == target_date:
            return col
    # すべて埋まっていれば末尾に追加
    new_col = ws.max_column + 1
    ws.cell(row=1, column=new_col).value = datetime.datetime(
        target_date.year, target_date.month, target_date.day
    )
    return new_col


def build_labs_preview(ws, results: list[dict], col_idx: int) -> list[dict]:
    """プレビュー用行リストを返す。信頼度 < 1.0 は黄色、マッチなしは赤フラグ付き。"""
    index = _build_test_name_index(ws)
    rows = []
    for item in results:
        name = item.get("name", "")
        value = item.get("value")
        unit = item.get("unit", "")
        match = _fuzzy_match(name, index)
        if match:
            matched_name, row_idx, conf = match
            rows.append({
                "AI抽出名":     name,
                "マッチした項目": matched_name,
                "値":           value,
                "単位":         unit,
                "信頼度":       round(conf, 2),
                "書き込み先":   f"行{row_idx} 列{col_idx}",
                "_row":         row_idx,
                "_col":         col_idx,
                "_write":       True,
            })
        else:
            rows.append({
                "AI抽出名":     name,
                "マッチした項目": "⚠️ 未マッチ",
                "値":           value,
                "単位":         unit,
                "信頼度":       0.0,
                "書き込み先":   "—",
                "_row":         None,
                "_col":         None,
                "_write":       False,
            })
    return rows


def write_labs_data(results: list[dict], target_date: datetime.date, excel_path: str,
                    sheet_prefix: str = "") -> list[dict]:
    """
    Labsシートに書き込む。信頼度 < FUZZY_CUTOFF のものはスキップ。

    Returns:
        プレビュー行リスト
    """
    writer = ExcelWriter(excel_path)
    wb = writer.load()
    ws = ExcelWriter.get_sheet(wb, sheet_prefix + ExcelWriter.SHEET_LABS)

    col_idx = _find_or_create_date_col(ws, target_date)
    preview = build_labs_preview(ws, results, col_idx)

    for row in preview:
        if row["_write"] and row["値"] is not None:
            ws.cell(row=row["_row"], column=row["_col"]).value = row["値"]

    writer.save(wb)
    return preview
