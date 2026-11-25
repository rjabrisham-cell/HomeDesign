# -*- coding: utf-8 -*-
"""
Apply manually-collected web person-ID hits for brand mobiles.
Also merge structured hits from a JSONL evidence file.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

ROOT = Path(r"D:\HomeDesign")
OUT = ROOT / "artifacts/Distribution-Contact-Directory-FA.xlsx"
LOCAL = ROOT / "research/distribution-contacts"
HITS = LOCAL / "batches/person_id_hits.jsonl"
YELLOW = PatternFill("solid", fgColor="FFF2CC")
BODY = Font(name="Vazirmatn", size=11)
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def to_mobile(v: str) -> str | None:
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    if len(d) == 11 and d.startswith("09"):
        return d
    return None


def main() -> None:
    if not HITS.exists():
        print("no hits file yet")
        return
    hits = []
    for line in HITS.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line:
            continue
        hits.append(json.loads(line))

    wb = openpyxl.load_workbook(OUT)
    ws = wb.active
    for h in hits:
        co = h.get("company") or h.get("brand") or "—"
        pe = h.get("person")
        mob = to_mobile(h.get("mobile", ""))
        if not pe or not mob:
            continue
        row = ws.max_row + 1
        ws.cell(row, 1, co).font = BODY
        ws.cell(row, 2, pe).font = BODY
        c = ws.cell(row, 3, mob)
        c.number_format = "@"
        c.font = BODY
        for col in (1, 2, 3):
            ws.cell(row, col).fill = YELLOW
    ws.auto_filter.ref = f"A1:C{ws.max_row}"
    wb.save(OUT)
    wb.save(LOCAL / "Distribution-Contact-Directory-FA.xlsx")
    print("applied_hits", len(hits), "rows_now", ws.max_row - 1)


if __name__ == "__main__":
    main()
