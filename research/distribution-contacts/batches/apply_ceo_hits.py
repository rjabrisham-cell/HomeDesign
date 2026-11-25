# -*- coding: utf-8 -*-
"""Apply CEO/manager hits from logiscm, APIR board, association list, and JSONL."""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill

ROOT = Path(r"D:\HomeDesign")
OUT = ROOT / "artifacts/Distribution-Contact-Directory-FA.xlsx"
LOCAL = ROOT / "research/distribution-contacts"
YELLOW = PatternFill("solid", fgColor="FFF2CC")
BODY = Font(name="Vazirmatn", size=11)
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def S(v) -> str:
    return "" if v is None else str(v).strip()


def to_mobile(v: str) -> str | None:
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
    t = re.sub(r"[()\[\]{}]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    for p in ("شرکت ", "شركت ", "پخش ", "گروه "):
        if t.startswith(p):
            t = t[len(p) :]
    return t


def has_person(pe: str) -> bool:
    pe = S(pe)
    return bool(pe) and not any(
        x in pe for x in ["ثبت نشده", "مرتبط با", "مدیر جدید", "ب نام شرکت"]
    )


def classify(co, pe, mobs):
    hp = has_person(pe)
    hc = bool(co) and co not in ("—", "-", "")
    hm = bool(mobs)
    if hp and hm and hc:
        return "triple"
    if hp and hc:
        return "pair_co_pe"
    if hc and hm:
        return "pair_co_mob"
    if hp and hm:
        return "pair_pe_mob"
    return "other"


def normalize_land(v: str) -> str | None:
    d = re.sub(r"\D", "", S(v).translate(DIGITS))
    if not d:
        return None
    if d.startswith("98"):
        d = "0" + d[2:]
    if d.startswith("09") or (len(d) == 10 and d.startswith("9")):
        return None
    if not d.startswith("0") and len(d) == 10:
        d = "0" + d
    elif not d.startswith("0") and len(d) == 8:
        d = "021" + d
    return d


def collect_hits() -> list[dict]:
    hits: list[dict] = []
    logiscm = json.loads((LOCAL / "queues/logiscm_rows.json").read_text(encoding="utf-8"))
    for row in logiscm:
        hits.append(
            {
                "company": row["company"],
                "person": f"{row['ceo']} — مدیرعامل (logiscm)",
                "mobile": None,
                "land": row.get("phone", ""),
                "src": "logiscm",
            }
        )

    board = [
        ("دراژه", "سهراب کارگر", "رئیس هیئت مدیره انجمن / مدیرعامل"),
        ("هلدینگ شفا", "محمدرضا مرادی", "نائب رئیس انجمن / مدیرعامل"),
        ("قاسم ایران", "حسین یوسف خزرائی", "خزانه‌دار انجمن / مدیرعامل"),
        ("آرتا مهر", "مجید جعفر نژاد", "عضو هیئت مدیره / مدیرعامل"),
        ("طلوع پخش آفتاب", "حمیدرضا صفرشریف", "عضو هیئت مدیره / رئیس هیئت مدیره شرکت"),
        ("سایه سمن", "مقصود نصیرزاده", "عضو هیئت مدیره / مدیرعامل"),
        ("گلپخش اول", "مجتبی کریمی", "عضو هیئت مدیره / مدیرعامل"),
        ("راسن درمان", "نادر علی زاد نمینی", "بازرس / مدیرعامل"),
        ("الوند یاران سحر", "محمود توتونچیان", "علی‌البدل هیئت مدیره / مدیرعامل"),
        ("پراور پخش سپاهان", "حسین همتی نژاد", "علی‌البدل هیئت مدیره / مدیرعامل"),
        ("کاسپین", "جبار صفی", "بازرس علی‌البدل / مدیرعامل"),
        ("تیپیکو", "عقیل آرین نژاد", "عضو هیئت مدیره"),
        ("مدبران کوشا ایرانیان", "مصطفی خدایی", "نائب رئیس سابق هیئت مدیره"),
        ("ساتراپ دینا", "محمد میربلوکی", "عضو هیئت مدیره"),
        ("پخش رازی", "محمدرضا مرادی", "خزانه‌دار / مدیرعامل"),
        ("پخش پدیده پایدار", "مجید جعفر نژاد", "عضو هیئت مدیره"),
        ("خاقانی نوین سپاهان", "جمشید سلطانی تهرانی", "رئیس هیئت مدیره انجمن اصفهان / مدیرعامل"),
        ("مهر آیین آریا ایرانیان", "مدیرعامل مهر آیین آریا ایرانیان", "عضو هیئت مدیره انجمن اصفهان"),
    ]
    for co, pe, role in board:
        hits.append(
            {
                "company": co,
                "person": f"{pe} — {role} (apir/apesf)",
                "mobile": None,
                "land": "",
                "src": "board",
            }
        )

    assoc = json.loads((LOCAL / "association_members_raw.json").read_text(encoding="utf-8"))
    for a in assoc:
        pe = S(a.get("ceo_source"))
        if pe:
            hits.append(
                {
                    "company": S(a.get("name_raw")),
                    "person": f"{pe} — مدیرعامل/مدیر (فهرست انجمن ملی)",
                    "mobile": None,
                    "land": S(a.get("phone_official")),
                    "src": "assoc",
                }
            )

    hits_path = LOCAL / "batches/person_id_hits.jsonl"
    if hits_path.exists():
        for line in hits_path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line:
                continue
            h = json.loads(line)
            hits.append(
                {
                    "company": h.get("company") or h.get("brand") or "—",
                    "person": h.get("person"),
                    "mobile": h.get("mobile"),
                    "land": h.get("land") or h.get("phone") or "",
                    "src": h.get("src") or "jsonl",
                }
            )
    return hits


def apply_hits(hits: list[dict]) -> dict:
    wb = openpyxl.load_workbook(OUT)
    ws = wb.active
    index: dict[str, list[int]] = defaultdict(list)
    for i, r in enumerate(ws.iter_rows(min_row=2, max_col=1), start=2):
        index[brand_key(S(r[0].value))].append(i)

    filled = 0
    added = 0
    mobile_fills = 0
    land_updates = 0

    for h in hits:
        co = S(h.get("company"))
        pe = S(h.get("person"))
        if not co or not pe:
            continue
        mob = to_mobile(h.get("mobile") or "")
        land = normalize_land(h.get("land") or "")
        k = brand_key(co)
        rows = index.get(k) or []

        def append_row(company: str, person: str, mobile: str | None, landline: str | None) -> int:
            row = ws.max_row + 1
            ws.cell(row, 1, company).font = BODY
            ws.cell(row, 2, person).font = BODY
            c = ws.cell(row, 3, mobile)
            c.number_format = "@"
            c.font = BODY
            d = ws.cell(row, 4, landline)
            d.number_format = "@"
            d.font = BODY
            for col in (1, 2, 3, 4):
                if ws.cell(row, col).value:
                    ws.cell(row, col).fill = YELLOW
            index[k].append(row)
            return row

        if rows:
            # Prefer filling blank person on an existing row that already has this mobile
            filled_here = False
            if mob:
                for i in rows:
                    cur_m = [
                        to_mobile(x)
                        for x in S(ws.cell(i, 3).value).splitlines()
                        if to_mobile(x)
                    ]
                    if mob in cur_m and not has_person(S(ws.cell(i, 2).value)):
                        ws.cell(i, 2).value = pe
                        ws.cell(i, 2).fill = YELLOW
                        ws.cell(i, 2).font = BODY
                        filled += 1
                        filled_here = True
                        break
            if not filled_here:
                # If any row has empty person and no conflicting identity, fill first blank
                for i in rows:
                    if not has_person(S(ws.cell(i, 2).value)):
                        ws.cell(i, 2).value = pe
                        ws.cell(i, 2).fill = YELLOW
                        ws.cell(i, 2).font = BODY
                        filled += 1
                        if mob:
                            cur_m = [
                                to_mobile(x)
                                for x in S(ws.cell(i, 3).value).splitlines()
                                if to_mobile(x)
                            ]
                            if mob not in cur_m:
                                cur_m.append(mob)
                                c = ws.cell(i, 3, "\n".join(cur_m))
                                c.number_format = "@"
                                c.fill = YELLOW
                                mobile_fills += 1
                        filled_here = True
                        break
            # Additional distinct person (+ optional mobile) → new accepted row
            if not filled_here:
                existing_people = {S(ws.cell(i, 2).value) for i in rows}
                existing_mobs: set[str] = set()
                for i in rows:
                    for x in S(ws.cell(i, 3).value).splitlines():
                        mm = to_mobile(x)
                        if mm:
                            existing_mobs.add(mm)
                person_new = pe not in existing_people
                mobile_new = bool(mob) and mob not in existing_mobs
                if person_new or mobile_new:
                    append_row(S(ws.cell(rows[0], 1).value) or co, pe, mob, land)
                    added += 1
            # still merge landlines onto first row
            if land:
                i = rows[0]
                cur_l = [x.strip() for x in S(ws.cell(i, 4).value).splitlines() if x.strip()]
                if land not in cur_l:
                    cur_l.append(land)
                    c = ws.cell(i, 4, "\n".join(cur_l))
                    c.number_format = "@"
                    land_updates += 1
        else:
            append_row(co, pe, mob, land)
            added += 1

    wb.save(OUT)
    wb.save(LOCAL / "Distribution-Contact-Directory-FA.xlsx")

    wb = openpyxl.load_workbook(OUT, data_only=True)
    counts: dict[str, int] = defaultdict(int)
    for r in wb.active.iter_rows(min_row=2, values_only=True):
        co, pe, mcol = S(r[0]), S(r[1]), S(r[2])
        mobs = [to_mobile(x) for x in mcol.splitlines() if to_mobile(x)] if mcol else []
        counts[classify(co, pe, mobs)] += 1
    acc = (
        counts["triple"]
        + counts["pair_co_pe"]
        + counts["pair_co_mob"]
        + counts["pair_pe_mob"]
    )
    stats = {
        "filled": filled,
        "added": added,
        "mobile_fills": mobile_fills,
        "land_updates": land_updates,
        "rows": wb.active.max_row - 1,
        **counts,
        "accepted_kpi": acc,
        "gap": max(0, 5000 - acc),
    }
    (LOCAL / "batches/pairs_triples_stats.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return stats


def main() -> None:
    hits = collect_hits()
    stats = apply_hits(hits)
    print(json.dumps(stats, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
