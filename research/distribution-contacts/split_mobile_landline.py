# -*- coding: utf-8 -*-
"""
Split phones into two columns:
  شرکت | فرد مؤثر | شماره موبایل | تلفن
Landlines are kept but excluded from accepted KPI.
Accepted KPI shapes:
  - company + person + mobile
  - company + person
  - company + mobile
  - person + mobile
"""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT = Path(r"D:\HomeDesign\artifacts\Distribution-Contact-Directory-FA.xlsx")
LOCAL = Path(r"D:\HomeDesign\research\distribution-contacts")
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Vazirmatn", bold=True, color="FFFFFF", size=12)
BODY = Font(name="Vazirmatn", size=11)
THIN = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)


def S(v) -> str:
    return "" if v is None else str(v).strip()


def to_mobile(v) -> str | None:
    d = re.sub(r"\D", "", S(v).translate(DIGITS))
    if not d:
        return None
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    if len(d) == 11 and d.startswith("09"):
        return d
    return None


def normalize_landline(v) -> str | None:
    raw = S(v).translate(DIGITS)
    if not raw:
        return None
    if to_mobile(raw):
        return None
    d = re.sub(r"\D", "", raw)
    if not d or len(d) < 5:
        # keep short switchboard codes as-is if original had digits
        if re.search(r"\d", raw):
            return raw.replace(" ", "")
        return None
    if d.startswith("98") and len(d) >= 12:
        return "0" + d[2:]
    if len(d) == 10 and d.startswith(("21", "26", "31", "41", "51", "61", "71", "81", "11", "13", "17", "34", "35", "44", "45", "54", "56", "58", "76", "77", "83", "84", "86", "87")):
        return "0" + d
    # preserve original formatting lightly
    return raw.replace(" ", "")


def split_phones(text: str) -> tuple[list[str], list[str]]:
    mobiles: list[str] = []
    lands: list[str] = []
    seen_m: set[str] = set()
    seen_l: set[str] = set()
    if not text:
        return mobiles, lands
    for line in re.split(r"[\n]+", S(text)):
        line = line.strip()
        if not line:
            continue
        # a line may contain multiple numbers
        parts = re.split(r"[,;/|]+", line)
        for part in parts:
            part = part.strip()
            if not part:
                continue
            m = to_mobile(part)
            if m:
                if m not in seen_m:
                    seen_m.add(m)
                    mobiles.append(m)
                continue
            land = normalize_landline(part)
            if land:
                key = re.sub(r"\D", "", land.translate(DIGITS)) or land
                if key not in seen_l:
                    seen_l.add(key)
                    lands.append(land)
    return mobiles, lands


def has_person(pe: str) -> bool:
    pe = S(pe)
    if not pe:
        return False
    return not any(x in pe for x in ("ثبت نشده", "مرتبط با", "مدیر جدید"))


def classify(co: str, pe: str, mobiles: list[str]) -> str:
    hp = has_person(pe)
    hc = bool(co) and co not in ("—", "-", "")
    hm = bool(mobiles)
    if hp and hm and hc:
        return "triple"
    if hp and hc:
        return "pair_co_pe"
    if hc and hm:
        return "pair_co_mob"
    if hp and hm:
        return "pair_pe_mob"
    if hc and False:
        return "unused"
    if hc:
        return "co_only_or_landline"
    return "weak"


def main() -> None:
    wb = openpyxl.load_workbook(OUT)
    ws = wb.active
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        co, pe = S(r[0]), S(r[1])
        # support old 3-col or already 4-col
        if ws.max_column >= 4:
            mob_text, land_text = S(r[2]), S(r[3])
            mobiles, lands = split_phones(mob_text)
            m2, l2 = split_phones(land_text)
            # if mobiles were wrongly in land col or vice versa, re-split union
            all_text = "\n".join(x for x in (mob_text, land_text) if x)
            mobiles, lands = split_phones(all_text)
        else:
            mobiles, lands = split_phones(S(r[2]))
        rows.append({"company": co, "person": pe, "mobiles": mobiles, "landlines": lands})

    counts = defaultdict(int)
    for r in rows:
        counts[classify(r["company"], r["person"], r["mobiles"])] += 1
    accepted = (
        counts["triple"]
        + counts["pair_co_pe"]
        + counts["pair_co_mob"]
        + counts["pair_pe_mob"]
    )

    nwb = openpyxl.Workbook()
    nws = nwb.active
    nws.title = "دفتر تماس"
    nws.sheet_view.rightToLeft = True
    nws.append(["شرکت", "فرد مؤثر", "شماره موبایل", "تلفن"])
    for rec in rows:
        nws.append(
            [
                rec["company"] or None,
                rec["person"] or None,
                "\n".join(rec["mobiles"]) if rec["mobiles"] else None,
                "\n".join(rec["landlines"]) if rec["landlines"] else None,
            ]
        )
        row = nws.max_row
        nws.cell(row, 3).number_format = "@"
        nws.cell(row, 4).number_format = "@"

    for cell in nws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in nws.iter_rows(min_row=2, max_row=nws.max_row, max_col=4):
        for cell in row:
            cell.font = BODY
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            cell.border = THIN
    nws.freeze_panes = "A2"
    nws.auto_filter.ref = f"A1:D{nws.max_row}"
    for col, width in zip("ABCD", (34, 40, 22, 22)):
        nws.column_dimensions[col].width = width

    nwb.save(OUT)
    nwb.save(LOCAL / "Distribution-Contact-Directory-FA.xlsx")

    stats = {
        "rows": len(rows),
        "triple_co_pe_mob": counts["triple"],
        "pair_co_pe": counts["pair_co_pe"],
        "pair_co_mob": counts["pair_co_mob"],
        "pair_pe_mob": counts["pair_pe_mob"],
        "other_rows": counts["co_only_or_landline"] + counts["weak"],
        "accepted_kpi": accepted,
        "target": 5000,
        "gap_accepted": max(0, 5000 - accepted),
        "rows_with_mobile": sum(1 for r in rows if r["mobiles"]),
        "rows_with_landline": sum(1 for r in rows if r["landlines"]),
        "columns": ["شرکت", "فرد مؤثر", "شماره موبایل", "تلفن"],
        "note": "تلفن ثابت در شمارش اهداف نیست؛ فقط نگهداری می‌شود.",
    }
    (LOCAL / "batches" / "pairs_triples_stats.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(stats, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
