# -*- coding: utf-8 -*-
"""
Rebuild directory maximizing accepted shapes:
- triple: company + person + mobile
- pair: company + person
- pair: company + mobile
- pair: person + mobile
Also keeps company + landline as secondary leads (not primary KPI).
"""
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
LOGISCM = Path(
    r"C:\Users\rjabr\.cursor\projects\D-HomeDesign\agent-tools\00dd9f90-9082-4a60-8eee-5e4d98f94e3f.txt"
)

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


def phones_from(*vals) -> list[str]:
    out: list[str] = []
    seen = set()
    for v in vals:
        if v is None or v == "":
            continue
        raw = S(v)
        # prefer mobile extraction per chunk
        chunks = re.split(r"[\n,;/|]+", raw)
        for ch in chunks + ([raw] if len(chunks) > 1 else []):
            m = to_mobile(ch)
            if m:
                if m not in seen:
                    seen.add(m)
                    out.append(m)
                continue
            d = re.sub(r"\D", "", ch.translate(DIGITS))
            if not d or len(d) < 5:
                continue
            # keep landline-ish text compact
            t = re.sub(r"\s+", "", ch.translate(DIGITS))
            if t and t not in seen and not to_mobile(t):
                # normalize leading 98 landlines loosely
                if d.startswith("98") and len(d) >= 12:
                    t = "0" + d[2:]
                elif len(d) == 10 and d.startswith("21"):
                    t = "0" + d
                seen.add(t)
                out.append(t)
    return out


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
    for p in ("آقای ", "اقای ", "خانم ", "مهندس ", "دکتر ", "آقاي ", "دكتر "):
        if t.startswith(p):
            t = t[len(p) :]
    return t.strip()


def has_person(pe: str) -> bool:
    pe = S(pe)
    if not pe:
        return False
    return not any(x in pe for x in ("ثبت نشده", "مرتبط با", "مدیر جدید"))


def classify(co: str, pe: str, phones: list[str]) -> str:
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
    records: list[dict] = []

    def add(co, pe, phones, uncertain=False, source=""):
        co = norm_co(co)
        pe = S(pe)
        phones = [p for p in phones if p]
        if not co and not pe and not phones:
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

    # --- prior excel ---
    if OUT.exists():
        ws = openpyxl.load_workbook(OUT).active
        for r in ws.iter_rows(min_row=2, values_only=True):
            phones = [x.strip() for x in S(r[2]).splitlines() if x.strip()] if r[2] else []
            add(r[0], r[1], phones, False, "prior")

    # --- brand files ---
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
                    person = pe if not role else f"{pe} — {role}"
                    phs = phones_from(r[2])
                    key = brand_key(brand)
                    if pe:
                        brand_people[key].add(person)
                        add(brand, person, phs, True, f"brand_report:{sn}")
                    for m in phs:
                        if to_mobile(m):
                            brand_mobiles[key].add(to_mobile(m))
                            mobile_brands[to_mobile(m)].add(brand)
            elif "برند" in sn or "همه" in sn:
                for r in rows[1:]:
                    if not r or len(r) < 3:
                        continue
                    brand = S(r[1])
                    key = brand_key(brand)
                    phs = []
                    for col in (2, 3, 5):
                        if col < len(r):
                            phs.extend(phones_from(r[col]))
                    if brand and phs:
                        add(brand, "", phs, True, f"brand_sheet:{sn}")
                    for m in phs:
                        mm = to_mobile(m)
                        if mm:
                            brand_mobiles[key].add(mm)
                            mobile_brands[mm].add(brand)

    # crosslink brand people <-> mobiles
    for key, peoples in brand_people.items():
        for pe in peoples:
            for m in brand_mobiles.get(key, ()):
                brand = next(iter(mobile_brands[m]))
                add(brand, pe, [m], True, "brand_crosslink")

    # --- workbook_full ---
    w = openpyxl.load_workbook(UP / "workbook_full_df20.xlsx", data_only=True)
    for r in w.active.iter_rows(min_row=2, values_only=True):
        co, pe, role = S(r[1]), S(r[4]), S(r[5])
        phs = phones_from(r[2], r[3])
        person = pe if not role else f"{pe} — {role}"
        if pe or phs or co:
            add(co, person if pe else "", phs, False, "workbook_full")

    # --- info companies queue ---
    w = openpyxl.load_workbook(UP / "_______________1438.xlsx", data_only=True)
    for r in w.active.iter_rows(min_row=2, values_only=True):
        co, city, phone, email, pe = S(r[0]), S(r[1]), r[2], r[3], S(r[4])
        phs = phones_from(phone)
        label = co
        if city and co:
            # city kept only in local evidence later; don't pollute company col
            pass
        if co or pe or phs:
            add(co, pe, phs, True, "info_queue")

    # --- association committee ---
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
        c_pe = col("نام مديرعامل", "نام مدیرعامل", "نام و نام خانوادگی")
        c_tel = col("تلفن تماس", "تلفن")
        c_mob = col("تلفن همراه")
        for r in rows[1:]:
            if not r:
                continue

            def g(c):
                return r[c] if c is not None and c < len(r) else None

            co, pe = S(g(c_co)), S(g(c_pe))
            phs = phones_from(g(c_mob), g(c_tel))
            if pe:
                add(co, f"{pe} — مدیرعامل", phs, False, f"assoc:{sn}")
            elif co:
                add(co, "", phs, False, f"assoc:{sn}")

    # sheet کمیته غذایی (2) people with roles
    if "کمیته غذایی (2)" in w.sheetnames:
        ws = w["کمیته غذایی (2)"]
        rows = list(ws.iter_rows(min_row=1, values_only=True))
        hdr = [S(h) for h in rows[0]]

        def col2(*names):
            for i, h in enumerate(hdr):
                for n in names:
                    if n in h:
                        return i
            return None

        c_co = col2("نام شركت", "نام شرکت")
        c_pe = col2("نام و نام خانوادگی")
        c_role = col2("سمت")
        c_tel = col2("تلفن تماس")
        c_mob = col2("تلفن همراه")
        for r in rows[1:]:
            if not r:
                continue

            def g(c):
                return r[c] if c is not None and c < len(r) else None

            co, pe, role = S(g(c_co)), S(g(c_pe)), S(g(c_role))
            phs = phones_from(g(c_mob), g(c_tel))
            if pe:
                person = pe if not role else f"{pe} — {role}"
                add(co, person, phs, False, "assoc:committee2")

    # --- Iran B2B directory ---
    iran = UP / "Iran_Distribution_B2B_Directory_736a.xlsx"
    if iran.exists():
        w = openpyxl.load_workbook(iran, data_only=True)
        if "01_Companies" in w.sheetnames:
            for r in w["01_Companies"].iter_rows(min_row=2, values_only=True):
                co = S(r[1]) or S(r[2])
                phs = phones_from(r[8])
                add(co, "", phs, False, "iran_b2b_co")
        if "02_DecisionMakers" in w.sheetnames:
            for r in w["02_DecisionMakers"].iter_rows(min_row=2, values_only=True):
                co, pe, role = S(r[1]), S(r[2]), S(r[3])
                phs = phones_from(r[4], r[5])
                person = pe if not role else f"{pe} — {role}"
                if pe or phs:
                    add(co, person if pe else "", phs, True, "iran_b2b_dm")

    # --- Customer1399 company names (co_only; weak unless later enriched) ---
    cust = UP / "Customer1399_d14f.xlsx"
    if cust.exists():
        w = openpyxl.load_workbook(cust, data_only=True)
        for sn in w.sheetnames:
            ws = w[sn]
            for i, r in enumerate(ws.iter_rows(min_row=1, values_only=True)):
                if not r:
                    continue
                # first col often company
                co = S(r[0])
                if i == 0 and ("نام" in co or "شرکت" in co or "برند" in co or not co):
                    continue
                if len(co) >= 3:
                    add(co, "", [], True, f"customer1399:{sn}")

    # --- logiscm directory dump ---
    if LOGISCM.exists():
        text = LOGISCM.read_text(encoding="utf-8", errors="ignore")
        # table rows like: | N) | name | address | city | phone | email | ceo |
        for m in re.finditer(
            r"\|\s*\d+[\)\.]?\s*\|\s*([^|]+)\|\s*([^|]*)\|\s*([^|]*)\|\s*([^|]*)\|\s*([^|]*)\|\s*([^|]*)\|",
            text,
        ):
            name, _addr, _city, phone, _email, ceo = [S(x) for x in m.groups()]
            if not name or name.lower() in ("نام شرکت", "name"):
                continue
            phs = phones_from(phone)
            pe = ceo
            if pe and pe.lower() in ("مدیر عامل", "مدیرعامل", "-"):
                pe = ""
            add(name, pe, phs, True, "logiscm")

    # --- APIR board hard-coded public roles (co+pe) ---
    apir_board = [
        ("دراژه", "سهراب کارگر — رئیس هیئت مدیره انجمن / مدیرعامل"),
        ("هلدینگ شفا", "محمدرضا مرادی — نایب رئیس انجمن"),
        ("قاسم ایران", "حسین یوسف خزرائی — عضو هیئت مدیره انجمن / مدیرعامل"),
        ("آرتا مهر", "مجید جعفر نژاد — عضو هیئت مدیره انجمن / مدیرعامل"),
        ("طلوع پخش آفتاب", "حمیدرضا صفرشریف — عضو هیئت مدیره انجمن"),
        ("سایه سمن", "مقصود نصیرزاده — عضو هیئت مدیره انجمن / مدیرعامل"),
        ("گلپخش اول", "مجتبی کریمی — عضو هیئت مدیره انجمن / مدیرعامل"),
        ("راسن درمان", "نادر علی زاد نمینی — مدیرعامل"),
        ("الوند یاران سحر", "محمود توتونچیان — عضو علی‌البدل انجمن / مدیرعامل"),
        ("پراور پخش سپاهان", "حسین همتی نژاد — عضو علی‌البدل انجمن / مدیرعامل"),
        ("کاسپین", "جبار صفی — بازرس علی‌البدل انجمن / مدیرعامل"),
    ]
    for co, pe in apir_board:
        add(co, pe, [], False, "apir_board")

    # merge
    merged: dict = {}
    for rec in records:
        pe, co, phones = rec["person"], rec["company"], rec["phones"]
        pk = person_key(pe) if pe else ""
        mobs = [to_mobile(p) for p in phones if to_mobile(p)]
        if pk and mobs:
            key = ("pm", pk, mobs[0])
        elif pk:
            key = ("p", pk, brand_key(co))
        elif mobs:
            key = ("m", brand_key(co), mobs[0])
        else:
            key = ("c", brand_key(co), tuple(sorted({re.sub(r'\D','',S(p).translate(DIGITS)) for p in phones if S(p)}) or ("",)))
        if key not in merged:
            merged[key] = dict(rec)
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
        if not has_person(old["person"]) and has_person(pe):
            old["person"] = pe
        if old["company"] in ("—", "") and co:
            old["company"] = co
        old["uncertain"] = old.get("uncertain", False) or rec["uncertain"]

    final = list(merged.values())

    counts = defaultdict(int)
    for r in final:
        counts[classify(r["company"], r["person"], r["phones"])] += 1
    accepted = (
        counts["triple"]
        + counts["pair_co_pe"]
        + counts["pair_co_mob"]
        + counts["pair_pe_mob"]
    )
    stats = {
        "rows": len(final),
        "triple": counts["triple"],
        "pair_co_pe": counts["pair_co_pe"],
        "pair_co_mob": counts["pair_co_mob"],
        "pair_pe_mob": counts["pair_pe_mob"],
        "co_phone": counts["co_phone"],
        "co_only": counts["co_only"],
        "weak": counts["weak"],
        "accepted_kpi": accepted,
        "target": 5000,
        "gap_accepted": max(0, 5000 - accepted),
        "accepted_plus_co_phone": accepted + counts["co_phone"],
    }

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
        shape = classify(rec["company"], rec["person"], rec["phones"])
        # yellow if uncertain OR incomplete pair (not triple)
        if rec.get("uncertain") or shape != "triple":
            if shape in ("triple", "pair_co_pe", "pair_co_mob", "pair_pe_mob"):
                # only mark uncertain incomplete web/brand links yellow
                if rec.get("uncertain"):
                    for c in range(1, 4):
                        if nws.cell(row, c).value:
                            nws.cell(row, c).fill = YELLOW
            elif rec.get("uncertain"):
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
    (LOCAL / "batches" / "pairs_triples_stats.json").write_text(
        json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(stats, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
