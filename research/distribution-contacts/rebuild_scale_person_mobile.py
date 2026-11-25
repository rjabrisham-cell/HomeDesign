# -*- coding: utf-8 -*-
"""Rebuild contact directory maximizing person+mobile; write person-ID queue."""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

UP = Path(r"C:\Users\rjabr\.cursor\projects\D-HomeDesign\uploads")
OUT = Path(r"D:\HomeDesign\artifacts\Distribution-Contact-Directory-FA.xlsx")
LOCAL = Path(r"D:\HomeDesign\research\distribution-contacts")
QUEUE = LOCAL / "queues"
QUEUE.mkdir(parents=True, exist_ok=True)

DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
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
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def to_mobile(v) -> str | None:
    raw = S(v)
    if not raw:
        return None
    for part in re.split(r"[\s,/|;\n\\]+", raw) + [raw]:
        d = re.sub(r"\D", "", part.translate(DIGITS))
        if not d:
            continue
        if d.startswith("98") and len(d) >= 12:
            d = "0" + d[2:]
        if len(d) == 10 and d.startswith("9"):
            d = "0" + d
        if len(d) == 11 and d.startswith("09"):
            return d
    return None


def norm_co(s: str) -> str:
    t = S(s).replace("ي", "ی").replace("ك", "ک")
    return re.sub(r"\s+", " ", t).strip()


def brand_key(s: str) -> str:
    t = norm_co(s).lower()
    for p in ("شرکت ", "شركت ", "پخش "):
        if t.startswith(p):
            t = t[len(p) :]
    return re.sub(r"\s+", " ", t).strip()


def person_key(s: str) -> str:
    t = re.sub(r"\s+", " ", S(s).split("—")[0].split("|")[0])
    for p in ("آقای ", "اقای ", "خانم ", "مهندس ", "دکتر "):
        if t.startswith(p):
            t = t[len(p) :]
    return t.strip()


def main() -> None:
    records: list[dict] = []

    def add_rec(co, pe, phones, uncertain=False, source=""):
        co = norm_co(co)
        pe = S(pe)
        phones = [p for p in phones if p]
        if not co and not pe:
            return
        records.append(
            {
                "company": co or "—",
                "person": pe,
                "phones": phones,
                "uncertain": uncertain,
                "source": source,
            }
        )

    # prior working file
    if OUT.exists():
        ws = openpyxl.load_workbook(OUT).active
        for r in ws.iter_rows(min_row=2, values_only=True):
            co, pe, ph = S(r[0]), S(r[1]), S(r[2])
            phones = [x.strip() for x in ph.splitlines() if x.strip()] if ph else []
            add_rec(co, pe, phones, False, "prior")

    brand_people: dict[str, set[str]] = defaultdict(set)
    brand_mobiles: dict[str, set[str]] = defaultdict(set)
    mobile_brands: dict[str, set[str]] = defaultdict(set)

    for fname in [
        "_____________________________9bb0.xlsx",
        "________________bbdb.xlsx",
    ]:
        w = openpyxl.load_workbook(UP / fname, data_only=True)
        for sn in w.sheetnames:
            rows = list(w[sn].iter_rows(min_row=1, values_only=True))
            if not rows:
                continue
            if "گزارش" in sn:
                for r in rows[1:]:
                    if not r or len(r) < 5:
                        continue
                    brand, pe = S(r[1]), S(r[4])
                    role = S(r[5]) if len(r) > 5 else ""
                    m = to_mobile(r[2])
                    person = pe if not role else f"{pe} — {role}"
                    key = brand_key(brand)
                    if pe:
                        brand_people[key].add(person)
                    if m:
                        brand_mobiles[key].add(m)
                        mobile_brands[m].add(brand)
                        if pe:
                            add_rec(brand, person, [m], True, f"brand_report:{sn}")
            elif "برند" in sn or "همه" in sn:
                for r in rows[1:]:
                    if not r or len(r) < 3:
                        continue
                    brand = S(r[1])
                    key = brand_key(brand)
                    for col in (2, 3, 5):
                        if col < len(r):
                            m = to_mobile(r[col])
                            if m and brand:
                                brand_mobiles[key].add(m)
                                mobile_brands[m].add(brand)

    for key, peoples in brand_people.items():
        for pe in peoples:
            for m in brand_mobiles.get(key, ()):
                brand = next(iter(mobile_brands[m]))
                add_rec(brand, pe, [m], True, "brand_crosslink")

    w = openpyxl.load_workbook(UP / "workbook_full_df20.xlsx", data_only=True)
    for r in w.active.iter_rows(min_row=2, values_only=True):
        co, pe, role = S(r[1]), S(r[4]), S(r[5])
        phones = []
        for col in (2, 3):
            m = to_mobile(r[col])
            if m:
                phones.append(m)
            else:
                land = S(r[col])
                if land:
                    phones.append(land)
        if pe or phones:
            person = pe if not role else f"{pe} — {role}"
            add_rec(co, person if pe else "", phones, False, "workbook_full")

    w = openpyxl.load_workbook(UP / "______________________________________84f8.xlsx", data_only=True)
    for r in w["کمیته غذایی"].iter_rows(min_row=2, values_only=True):
        co, pe, tel = S(r[1]), S(r[2]), S(r[4])
        if pe:
            add_rec(co, f"{pe} — مدیرعامل", [tel] if tel else [], False, "assoc_committee")

    # brand mobile without person yet — keep as company-level mobile lead (not counted as person+mobile)
    for m, brands in mobile_brands.items():
        brand = sorted(brands)[0]
        add_rec(brand, "", [m], True, "brand_mobile_only")

    merged: dict = {}
    for rec in records:
        pe, co, phones = rec["person"], rec["company"], rec["phones"]
        pk = person_key(pe) if pe else ""
        mobs_m = [to_mobile(p) for p in phones if to_mobile(p)]
        if pk and mobs_m:
            key = ("pm", pk, mobs_m[0])
        elif pk:
            key = ("p", pk, brand_key(co))
        else:
            key = ("c", brand_key(co), tuple(sorted(set(mobs_m or [S(p) for p in phones]))))
        if key not in merged:
            merged[key] = rec
            continue
        old = merged[key]
        seen = set()
        out = []
        for p in old["phones"] + phones:
            d = re.sub(r"\D", "", S(p).translate(DIGITS))
            if not d or d in seen:
                continue
            seen.add(d)
            out.append(to_mobile(p) or p)
        old["phones"] = out
        if not old["person"] and pe:
            old["person"] = pe
        if old["company"] in ("—", "") and co:
            old["company"] = co
        old["uncertain"] = old["uncertain"] or rec["uncertain"]

    final = list(merged.values())
    pm = sum(1 for r in final if r["person"] and any(to_mobile(p) for p in r["phones"]))
    print("final_rows", len(final))
    print("person_mobile_rows", pm)
    print("target_gap", max(0, 5000 - pm))

    nwb = openpyxl.Workbook()
    nws = nwb.active
    nws.title = "دفتر تماس"
    nws.sheet_view.rightToLeft = True
    nws.append(["شرکت", "فرد مؤثر", "شماره موبایل و تلفن"])
    for rec in sorted(final, key=lambda x: (x["company"], x["person"])):
        ph = "\n".join(rec["phones"]) if rec["phones"] else None
        nws.append([rec["company"], rec["person"] or None, ph])
        row = nws.max_row
        nws.cell(row, 3).number_format = "@"
        if rec["uncertain"]:
            for c in range(1, 4):
                if nws.cell(row, c).value:
                    nws.cell(row, c).fill = YELLOW
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

    have = set()
    for r in final:
        if not r["person"]:
            continue
        for p in r["phones"]:
            m = to_mobile(p)
            if m:
                have.add(m)
    need = [
        {"mobile": m, "brands": sorted(brands)[:5]}
        for m, brands in mobile_brands.items()
        if m not in have
    ]
    (QUEUE / "mobiles_need_person.json").write_text(
        json.dumps(need, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    stats = {
        "final_rows": len(final),
        "person_mobile_rows": pm,
        "target": 5000,
        "gap": max(0, 5000 - pm),
        "mobiles_need_person": len(need),
        "brands_with_people": len(brand_people),
        "brands_with_mobiles": len(brand_mobiles),
    }
    (LOCAL / "batches" / "scale_stats.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(stats, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
