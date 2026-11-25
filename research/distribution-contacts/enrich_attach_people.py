# -*- coding: utf-8 -*-
"""Attach known CEOs/people to company rows missing person; recount accepted KPI."""
from __future__ import annotations

import json
import re
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

ROOT = Path(r"D:\HomeDesign")
OUT = ROOT / "artifacts/Distribution-Contact-Directory-FA.xlsx"
LOCAL = ROOT / "research/distribution-contacts"
LOGISCM = LOCAL / "queues/logiscm_rows.json"
UP = Path(r"C:\Users\rjabr\.cursor\projects\D-HomeDesign\uploads")

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
    t = re.sub(r"[\(\)\[\]\{\}«»\"']+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def has_person(pe: str) -> bool:
    pe = S(pe)
    if not pe:
        return False
    return not any(x in pe for x in ("ثبت نشده", "مرتبط با", "مدیر جدید"))


def phones_from(v) -> list[str]:
    out = []
    seen = set()
    raw = S(v).translate(DIGITS)
    if not raw:
        return out
    # logiscm style 9821...+
    d = re.sub(r"\D", "", raw)
    if d.startswith("98") and len(d) >= 12:
        cand = "0" + d[2:]
        m = to_mobile(cand)
        if m:
            return [m]
        out.append(cand)
        return out
    for part in re.split(r"[\n,;/|]+", raw):
        m = to_mobile(part)
        if m and m not in seen:
            seen.add(m)
            out.append(m)
            continue
        dd = re.sub(r"\D", "", part)
        if len(dd) >= 5 and dd not in seen:
            seen.add(dd)
            if dd.startswith("98") and len(dd) >= 12:
                out.append("0" + dd[2:])
            elif len(dd) == 10 and dd.startswith("21"):
                out.append("0" + dd)
            else:
                out.append(part.strip())
    return out


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
    # people map from logiscm + association
    people_map: dict[str, str] = {}
    phone_map: dict[str, list[str]] = defaultdict(list)

    if LOGISCM.exists():
        for row in json.loads(LOGISCM.read_text(encoding="utf-8")):
            co = S(row.get("company"))
            ceo = S(row.get("ceo"))
            phs = phones_from(row.get("phone"))
            k = brand_key(co)
            if ceo and k:
                people_map[k] = f"{ceo} — مدیرعامل/سرنخ فهرست صنفی"
            if phs and k:
                phone_map[k].extend(phs)

    # association
    w = openpyxl.load_workbook(UP / "______________________________________84f8.xlsx", data_only=True)
    for sn in ("کمیته غذایی", "لیست اعضا-1404"):
        if sn not in w.sheetnames:
            continue
        ws = w[sn]
        rows = list(ws.iter_rows(min_row=1, values_only=True))
        hdr = [S(h) for h in rows[0]]

        def col(*names):
            for i, h in enumerate(hdr):
                for n in names:
                    if n in h:
                        return i
            return None

        c_co = col("نام شركت", "نام شرکت")
        c_pe = col("نام مديرعامل", "نام مدیرعامل")
        c_tel = col("تلفن تماس", "تلفن")
        for r in rows[1:]:
            if not r:
                continue
            co = S(r[c_co]) if c_co is not None and c_co < len(r) else ""
            pe = S(r[c_pe]) if c_pe is not None and c_pe < len(r) else ""
            tel = r[c_tel] if c_tel is not None and c_tel < len(r) else None
            k = brand_key(co)
            if pe and k and k not in people_map:
                people_map[k] = f"{pe} — مدیرعامل"
            for p in phones_from(tel):
                phone_map[k].append(p)

    wb = openpyxl.load_workbook(OUT)
    ws = wb.active
    attached_people = 0
    attached_phones = 0
    existing_keys = set()

    for i, r in enumerate(ws.iter_rows(min_row=2, max_col=3), start=2):
        co = S(r[0].value)
        pe = S(r[1].value)
        ph = S(r[2].value)
        k = brand_key(co)
        existing_keys.add(k)
        phones = [x.strip() for x in ph.splitlines() if x.strip()] if ph else []
        if not has_person(pe) and k in people_map:
            r[1].value = people_map[k]
            r[1].fill = YELLOW
            attached_people += 1
        # add missing phones from map
        if k in phone_map:
            seen = {re.sub(r"\D", "", S(p).translate(DIGITS)) for p in phones}
            added = False
            for p in phone_map[k]:
                d = re.sub(r"\D", "", S(p).translate(DIGITS))
                if d and d not in seen:
                    phones.append(p)
                    seen.add(d)
                    added = True
            if added:
                r[2].value = "\n".join(phones)
                r[2].number_format = "@"
                r[2].fill = YELLOW
                attached_phones += 1

    # append logiscm companies not already present
    appended = 0
    for k, pe in people_map.items():
        if k in existing_keys:
            continue
        # find a display company name from logiscm
        co = k
        for row in json.loads(LOGISCM.read_text(encoding="utf-8")):
            if brand_key(row["company"]) == k:
                co = row["company"]
                break
        row = ws.max_row + 1
        ws.cell(row, 1, co).font = BODY
        ws.cell(row, 2, pe).font = BODY
        phs = phone_map.get(k, [])
        c = ws.cell(row, 3, "\n".join(phs) if phs else None)
        c.number_format = "@"
        c.font = BODY
        for col in (1, 2, 3):
            if ws.cell(row, col).value:
                ws.cell(row, col).fill = YELLOW
        appended += 1
        existing_keys.add(k)

    # restyle lightly
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:C{ws.max_row}"
    for col, width in zip("ABC", (36, 42, 28)):
        ws.column_dimensions[col].width = width
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    wb.save(OUT)
    wb.save(LOCAL / "Distribution-Contact-Directory-FA.xlsx")

    # recount
    wb2 = openpyxl.load_workbook(OUT, data_only=True)
    counts = defaultdict(int)
    for r in wb2.active.iter_rows(min_row=2, values_only=True):
        co, pe, ph = S(r[0]), S(r[1]), S(r[2])
        phones = [x.strip() for x in ph.splitlines() if x.strip()] if ph else []
        counts[classify(co, pe, phones)] += 1
    accepted = (
        counts["triple"]
        + counts["pair_co_pe"]
        + counts["pair_co_mob"]
        + counts["pair_pe_mob"]
    )
    stats = {
        "attached_people": attached_people,
        "attached_phones": attached_phones,
        "appended_companies": appended,
        "rows": wb2.active.max_row - 1,
        "triple": counts["triple"],
        "pair_co_pe": counts["pair_co_pe"],
        "pair_co_mob": counts["pair_co_mob"],
        "pair_pe_mob": counts["pair_pe_mob"],
        "co_phone": counts["co_phone"],
        "co_only": counts["co_only"],
        "accepted_kpi": accepted,
        "target": 5000,
        "gap_accepted": max(0, 5000 - accepted),
        "people_map_size": len(people_map),
    }
    (LOCAL / "batches" / "pairs_triples_stats.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(stats, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
