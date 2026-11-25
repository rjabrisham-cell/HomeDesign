# -*- coding: utf-8 -*-
"""Merge same-company rows so co+pe and co+mob combine into triples where possible."""
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
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    if len(d) == 11 and d.startswith("09"):
        return d
    return None


def brand_key(s: str) -> str:
    t = S(s).replace("ي", "ی").replace("ك", "ک").lower()
    t = re.sub(r"\s+", " ", t).strip()
    for p in ("شرکت ", "شركت ", "پخش ", "گروه "):
        if t.startswith(p):
            t = t[len(p) :]
    t = re.sub(r"[()\[\]{}«»\"']+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def has_person(pe: str) -> bool:
    pe = S(pe)
    if not pe:
        return False
    return not any(x in pe for x in ("ثبت نشده", "مرتبط با", "مدیر جدید"))


def person_key(s: str) -> str:
    t = re.sub(r"\s+", " ", S(s).split("—")[0].split("|")[0])
    for p in ("آقای ", "اقای ", "خانم ", "مهندس ", "دکتر ", "آقاي ", "دكتر "):
        if t.startswith(p):
            t = t[len(p) :]
    return t.strip()


def classify(co, pe, phones):
    hp = has_person(pe)
    mobs = [p for p in phones if to_mobile(p)]
    hc = bool(co) and co not in ("—", "-", "")
    if hp and mobs and hc:
        return "triple"
    if hp and hc:
        return "pair_co_pe"
    if hc and mobs:
        return "pair_co_mob"
    if hp and mobs:
        return "pair_pe_mob"
    if hc and phones:
        return "co_phone"
    if hc:
        return "co_only"
    return "weak"


def main() -> None:
    wb = openpyxl.load_workbook(OUT)
    ws = wb.active
    groups: dict[str, list] = defaultdict(list)
    for r in ws.iter_rows(min_row=2, values_only=True):
        co, pe, ph = S(r[0]), S(r[1]), S(r[2])
        phones = [x.strip() for x in ph.splitlines() if x.strip()] if ph else []
        groups[brand_key(co) or co].append({"company": co, "person": pe, "phones": phones})

    final = []
    for _k, items in groups.items():
        co = max((it["company"] for it in items), key=len)
        phones_all = []
        seen_ph = set()
        pe_map = {}
        for it in items:
            for p in it["phones"]:
                d = re.sub(r"\D", "", S(p).translate(DIGITS))
                if d and d not in seen_ph:
                    seen_ph.add(d)
                    phones_all.append(to_mobile(p) or p)
            if has_person(it["person"]):
                pk = person_key(it["person"])
                if pk not in pe_map or len(it["person"]) > len(pe_map[pk]):
                    pe_map[pk] = it["person"]
        people = list(pe_map.values())
        mobs = [p for p in phones_all if to_mobile(p)]
        lands = [p for p in phones_all if not to_mobile(p)]
        if people:
            for pe in people:
                final.append({"company": co, "person": pe, "phones": mobs + lands})
        else:
            final.append({"company": co, "person": "", "phones": mobs + lands})

    counts = defaultdict(int)
    for r in final:
        counts[classify(r["company"], r["person"], r["phones"])] += 1
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
    nws.append(["شرکت", "فرد مؤثر", "شماره موبایل و تلفن"])
    for rec in sorted(final, key=lambda x: (x["company"], x["person"])):
        ph = "\n".join(rec["phones"]) if rec["phones"] else None
        nws.append([rec["company"], rec["person"] or None, ph])
        nws.cell(nws.max_row, 3).number_format = "@"
    for cell in nws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in nws.iter_rows(min_row=2, max_row=nws.max_row, max_col=3):
        for cell in row:
            cell.font = BODY
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            cell.border = THIN
    nws.freeze_panes = "A2"
    nws.auto_filter.ref = f"A1:C{nws.max_row}"
    for col, width in zip("ABC", (36, 42, 28)):
        nws.column_dimensions[col].width = width
    nwb.save(OUT)
    nwb.save(LOCAL / "Distribution-Contact-Directory-FA.xlsx")

    stats = {
        "rows": len(final),
        "triple": counts["triple"],
        "pair_co_pe": counts["pair_co_pe"],
        "pair_co_mob": counts["pair_co_mob"],
        "pair_pe_mob": counts["pair_pe_mob"],
        "co_phone": counts["co_phone"],
        "co_only": counts["co_only"],
        "accepted_kpi": accepted,
        "target": 5000,
        "gap_accepted": max(0, 5000 - accepted),
    }
    (LOCAL / "batches" / "pairs_triples_stats.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(stats, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
