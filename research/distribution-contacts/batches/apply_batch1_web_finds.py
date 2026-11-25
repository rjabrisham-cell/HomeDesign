# -*- coding: utf-8 -*-
"""Apply batch-1 external web findings to Distribution-Contact-Directory-FA.xlsx
and write local evidence. Delta vs baseline counted by digit-only phone identity.
"""
from __future__ import annotations

import json
import re
import shutil
from datetime import datetime, timezone
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side

ROOT = Path(r"D:\HomeDesign")
BASELINE = ROOT / "artifacts/contact-directory-baseline/Distribution-Contact-Directory-FA.baseline.xlsx"
WORKING = ROOT / "artifacts/Distribution-Contact-Directory-FA.xlsx"
LOCAL = ROOT / "research/distribution-contacts"
EVID = LOCAL / "batches/batch1_evidence.json"
META = LOCAL / "batches/batch1_delta.json"

YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Vazirmatn", bold=True, color="FFFFFF", size=12)
BODY_FONT = Font(name="Vazirmatn", size=11)
THIN = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def phone_digits(s: str) -> str:
    return re.sub(r"\D", "", str(s).translate(DIGITS))


def phone_keys(s: str) -> set[str]:
    d = phone_digits(s)
    if not d:
        return set()
    keys = {d}
    d2 = d.lstrip("0")
    if d2:
        keys.add(d2)
    if len(d) >= 10:
        keys.add(d[-10:])
    if len(d) >= 8:
        keys.add(d[-8:])
    # strip Iran country 98
    if d.startswith("98") and len(d) >= 12:
        keys.add(d[2:])
        keys.add("0" + d[2:])
    return {k for k in keys if len(k) >= 7}


def collect_phone_universe(ws) -> set[str]:
    u: set[str] = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        ph = r[2]
        if not ph:
            continue
        for line in str(ph).splitlines():
            u |= phone_keys(line)
    return u


def collect_person_names(ws) -> set[str]:
    names = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        pe = (r[1] or "").strip()
        if pe:
            names.add(re.sub(r"\s+", " ", pe.split("—")[0].split("-")[0].strip()))
    return names


def style_sheet(ws):
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:C{ws.max_row}"
    for col, w in zip("ABC", (36, 42, 28)):
        ws.column_dimensions[col].width = w
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=3):
        for cell in row:
            cell.font = BODY_FONT
            cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
            cell.border = THIN
            if isinstance(cell.value, str) and cell.column == 3:
                cell.number_format = "@"


def find_rows(ws, company_substr: str):
    hits = []
    for i, r in enumerate(ws.iter_rows(min_row=2, max_col=3), start=2):
        co = str(r[0].value or "")
        if company_substr in co:
            hits.append(i)
    return hits


def set_phones(cell, phones: list[str], uncertain=False):
    existing = []
    if cell.value:
        existing = [x.strip() for x in str(cell.value).splitlines() if x.strip()]
    seen = set()
    merged = []
    for p in existing + phones:
        k = phone_digits(p)
        if not k or k in seen:
            continue
        # drop garbled mega-concat (>14 digits pure)
        if len(k) > 14 and re.fullmatch(r"\d+", k):
            continue
        seen.add(k)
        merged.append(p)
    cell.value = "\n".join(merged) if merged else None
    cell.number_format = "@"
    if uncertain:
        cell.fill = YELLOW


def set_person(cell, person: str, uncertain=False):
    if not person:
        return
    cur = (cell.value or "").strip()
    if not cur or cur.startswith("مدیر جدید") or "ثبت نشده" in cur:
        cell.value = person
    elif person.split("—")[0].strip() not in cur:
        # keep both as clue
        cell.value = cur + " | " + person
        cell.fill = YELLOW
        return
    if uncertain:
        cell.fill = YELLOW


def append_row(ws, company, person, phones, uncertain=False):
    row = ws.max_row + 1
    ws.cell(row, 1, company).font = BODY_FONT
    ws.cell(row, 2, person or None).font = BODY_FONT
    ph = "\n".join(phones) if phones else None
    c = ws.cell(row, 3, ph)
    c.number_format = "@"
    c.font = BODY_FONT
    if uncertain:
        for col in (1, 2, 3):
            if ws.cell(row, col).value:
                ws.cell(row, col).fill = YELLOW
    return row


def main():
    shutil.copy2(BASELINE, WORKING)
    wb_base = openpyxl.load_workbook(BASELINE, data_only=True)
    ws_base = wb_base.active
    base_phones = collect_phone_universe(ws_base)
    base_people = collect_person_names(ws_base)

    wb = openpyxl.load_workbook(WORKING)
    ws = wb.active

    evidence = []
    stats = {
        "batch": 1,
        "cases_researched": 0,
        "new_phones": 0,
        "new_people": 0,
        "job_moves": 0,
        "confirmed_prior": 0,
        "no_result": 0,
        "access_barrier": 0,
        "new_phone_values": [],
        "new_people_values": [],
        "moves": [],
    }

    def count_new_phones(phones: list[str]) -> list[str]:
        fresh = []
        for p in phones:
            keys = phone_keys(p)
            # new if no key overlaps baseline universe
            if keys and not (keys & base_phones):
                fresh.append(p)
        return fresh

    # --- Case updates ---
    cases = []

    # 1) پدیده پایدار — NEW landlines from irangaah; CEO clue
    rows = find_rows(ws, "پدیده پایدار")
    phones = ["02158124520", "02158124000"]
    fresh = count_new_phones(phones)
    for ri in rows:
        set_phones(ws.cell(ri, 3), phones, uncertain=True)
        set_person(ws.cell(ri, 2), "صادق بیان — مدیرعامل/سرنخ فهرست صنفی", uncertain=True)
    # new person راجی from Golrang LinkedIn exhibition post
    if not any("راجی" in (n or "") for n in base_people):
        append_row(
            ws,
            "شرکت پخش پدیده پایدار",
            "راجی — مدیرعامل/عضو هیئت مدیره (سرنخ نمایشگاه صنعت پخش / سیمرغ گلرنگ)",
            [],
            uncertain=True,
        )
        stats["new_people"] += 1
        stats["new_people_values"].append("راجی @ پدیده پایدار")
    stats["new_phones"] += len(fresh)
    stats["new_phone_values"].extend(fresh)
    cases.append(
        {
            "case": "پخش پدیده پایدار",
            "queries": [
                "پخش پدیده پایدار تلفن تهران",
                "site:irangaah.ir پخش پدیده پایدار",
                "t.me golrangfamily پخش پدیده پایدار",
                "راجی مدیرعامل پخش پدیده",
            ],
            "sources": [
                "https://irangaah.ir/business/services/tehran/pakhsh-padideh-paydar",
                "https://t.me/s/golrangfamily?q=%23%D9%BE%D8%AE%D8%B4_%D9%BE%D8%AF%DB%8C%D8%AF%D9%87_%D9%BE%D8%A7%DB%8C%D8%AF%D8%A7%D8%B1",
                "https://logiscm.ir/distribution/",
                "https://linkedin.com/company/simorghgolrang",
            ],
            "prior": "ردیف مبنا شمارهٔ به‌هم‌چسبیدهٔ ۸۸۷۸۶۸۶۴۸۸۷۸۱۵۲۸ و «مدیر جدید ثبت نشده»",
            "found": {"phones": phones, "people": ["صادق بیان", "راجی"]},
            "new_phones_vs_baseline": fresh,
            "note": "۰۲۱۸۸۷۸۶۸۶۴ در مبنا به‌صورت به‌هم‌چسبیده هست → شمارهٔ تازه حساب نشد؛ ۵۸۱۲۴۵۲۰/۵۸۱۲۴۰۰۰ تازه است.",
        }
    )

    # 2) گلرنگ پخش — confirm Katebi; add Hosseini + Kangavari
    rows = find_rows(ws, "گلرنگ پخش")
    for ri in rows:
        co = str(ws.cell(ri, 1).value or "")
        # fill empty duplicate with existing HQ phone (move/confirm, not new)
        if not ws.cell(ri, 3).value and co.strip() in ("گلرنگ پخش", "شرکت گلرنگ پخش"):
            set_phones(ws.cell(ri, 3), ["02184237000"], uncertain=False)
            stats["confirmed_prior"] += 1
        pe = str(ws.cell(ri, 2).value or "")
        if "کاتبی" in pe:
            stats["confirmed_prior"] += 1
    # new people
    for person, label in [
        ("سیدعابدین حسینی — قائم‌مقام مدیرعامل / عضو هیئت مدیره", "سیدعابدین حسینی"),
        ("کنگاوری — مدیر منطقه (بازدید شعبه یزد دی ۱۴۰۴)", "کنگاوری"),
    ]:
        if not any(label in (n or "") for n in base_people):
            append_row(ws, "شرکت گلرنگ پخش", person, [], uncertain=True)
            stats["new_people"] += 1
            stats["new_people_values"].append(f"{label} @ گلرنگ پخش")
    # extra HQ line from logiscm (84237513) — NEW digit sequence
    gp_fresh = count_new_phones(["02184237513"])
    for ri in find_rows(ws, "شرکت گلرنگ پخش"):
        set_phones(ws.cell(ri, 3), ["02184237513"], uncertain=True)
        break
    stats["new_phones"] += len(gp_fresh)
    stats["new_phone_values"].extend(gp_fresh)
    cases.append(
        {
            "case": "گلرنگ پخش",
            "queries": [
                "پیمان کاتبی گلرنگ پخش مدیرعامل",
                "سیدعابدین حسینی گلرنگ پخش قائم مقام",
                "linkedin golrangpakhsh بازدید یزد",
            ],
            "sources": [
                "https://www.linkedin.com/posts/golrangpakhsh_%D8%A8%D8%A7%D8%B2%D8%AF%DB%8C%D8%AF-%D9%85%D8%AF%DB%8C%D8%B1%DB%8C%D8%AA-%D9%85%D8%AD%D8%AA%D8%B1%D9%85-%D8%B9%D8%A7%D9%85%D9%84-%D8%A7%D8%B2-%D8%B4%D8%B9%D8%A8%D9%87-%DB%8C%D8%B2%D8%AF-activity-7470730054399004672-w9ZR",
                "https://rasmio.com/person/33D85E8/%D9%BE%DB%8C%D9%85%D8%A7%D9%86%20%DA%A9%D8%A7%D8%AA%D8%A8%DB%8C/",
                "https://logiscm.ir/distribution/",
            ],
            "prior": "پیمان کاتبی + 84237000 در مبنا",
            "found": {
                "phones": ["02184237513"],
                "people": ["سیدعابدین حسینی", "کنگاوری"],
                "confirmed": ["پیمان کاتبی مدیرعامل تا ۱۴۰۶/۰۳"],
            },
            "new_phones_vs_baseline": gp_fresh,
        }
    )

    # 3) پگاه دلپذیر / پخش پگاه
    phones = ["02161939000", "02161939173", "02177242308", "02136656620"]
    fresh = count_new_phones(phones)
    for ri in find_rows(ws, "پگاه دلپذیر"):
        set_phones(ws.cell(ri, 3), phones, uncertain=True)
        set_person(ws.cell(ri, 2), "مرتبط با شرکت پخش پگاه / برند دلپذیر", uncertain=True)
    # also company-level پخش پگاه if present
    if not find_rows(ws, "شرکت پخش پگاه") and not find_rows(ws, "پخش پگاه"):
        append_row(
            ws,
            "شرکت پخش پگاه",
            "مجید اخوان — مدیرعامل/سرنخ فهرست صنفی (املای اکحوان در منبع)",
            phones,
            uncertain=True,
        )
        stats["new_people"] += 1
        stats["new_people_values"].append("مجید اخوان @ پخش پگاه")
    else:
        for ri in find_rows(ws, "پخش پگاه"):
            set_phones(ws.cell(ri, 3), phones, uncertain=True)
            set_person(ws.cell(ri, 2), "مجید اخوان — مدیرعامل/سرنخ فهرست صنفی", uncertain=True)
    stats["new_phones"] += len(fresh)
    stats["new_phone_values"].extend(fresh)
    cases.append(
        {
            "case": "پخش پگاه دلپذیر",
            "queries": [
                "پخش پگاه دلپذیر تلفن",
                "مراکز پخش پگاه تهران",
                "دلپذیر تماس دفتر مرکزی",
            ],
            "sources": [
                "https://pegah.com/%d9%85%d8%b1%d8%a7%da%a9%d8%b2-%d9%be%d8%ae%d8%b4/",
                "https://pegah.com/%d8%aa%d9%85%d8%a7%d8%b3-%d8%a8%d8%a7-%d9%85%d8%a7/",
                "https://t.me/s/estekhdam_tehran/93212",
                "https://logiscm.ir/distribution/",
            ],
            "prior": "ردیف «پخش پگاه دلپذیر» بدون شماره",
            "found": {"phones": phones, "people": ["مجید اخوان"]},
            "new_phones_vs_baseline": fresh,
            "note": "آگهی مرداد ۱۴۰۵ رابطهٔ برند را نشان می‌دهد؛ شماره از سایت پگاه.",
        }
    )

    # 4) باران پخش پویا
    phones = ["02632260572", "02632260573", "02632260574"]
    fresh = count_new_phones(phones)
    for ri in find_rows(ws, "باران پخش پویا"):
        set_phones(ws.cell(ri, 3), phones, uncertain=True)
        set_person(ws.cell(ri, 2), "نادر پور محمد — مسئول شرکت/سرنخ foodkeys", uncertain=True)
    if "نادر پور محمد" not in " ".join(base_people):
        stats["new_people"] += 1
        stats["new_people_values"].append("نادر پور محمد @ باران پخش پویا")
    stats["new_phones"] += len(fresh)
    stats["new_phone_values"].extend(fresh)
    cases.append(
        {
            "case": "باران پخش پویا",
            "queries": ["باران پخش پویا تلفن کرج", "باران پخش پویا نادر پور محمد"],
            "sources": [
                "https://www.foodkeys.com/view/wholesale_trade/details?CompanyName=%D8%B4%D8%B1%DA%A9%D8%AA+%D8%A8%D8%A7%D8%B1%D8%A7%D9%86+%D9%BE%D8%AE%D8%B4+%D9%BE%D9%88%DB%8C%D8%A7&id=581"
            ],
            "prior": "بدون شماره در مبنا",
            "found": {"phones": phones, "people": ["نادر پور محمد"]},
            "new_phones_vs_baseline": fresh,
            "note": "منبع foodkeys هشدار به‌روز‌نبودن دارد → زرد.",
        }
    )

    # 5) آریان لیان پخش
    phones = ["02189328121"]
    fresh = count_new_phones(phones)
    for ri in find_rows(ws, "آریان لیان") + find_rows(ws, "اریان لیان"):
        set_phones(ws.cell(ri, 3), phones, uncertain=True)
    stats["new_phones"] += len(fresh)
    stats["new_phone_values"].extend(fresh)
    cases.append(
        {
            "case": "آریان لیان پخش",
            "queries": ["آریان لیان پخش تلفن", "alppdc.ir"],
            "sources": ["https://logiscm.ir/distribution/"],
            "prior": "بدون شماره",
            "found": {"phones": phones},
            "new_phones_vs_baseline": fresh,
        }
    )

    # 6) ایپک پخش یاران — phones from «پخش سراسری ایپک» ads; uncertain match
    phones = ["02144220363", "02144256854", "02144256853"]
    fresh = count_new_phones(phones)
    for ri in find_rows(ws, "ایپک پخش"):
        set_phones(ws.cell(ri, 3), phones, uncertain=True)
    stats["new_phones"] += len(fresh)
    stats["new_phone_values"].extend(fresh)
    cases.append(
        {
            "case": "ایپک پخش یاران",
            "queries": ["ایپک پخش یاران تلفن", "پخش سراسری ایپک استخدام"],
            "sources": [
                "https://www.momtaznews.com/%D8%A7%D8%B3%D8%AA%D8%AE%D8%AF%D8%A7%D9%85-%D8%B4%D8%B1%DA%A9%D8%AA-%D9%BE%D8%AE%D8%B4-%D8%B3%D8%B1%D8%A7%D8%B3%D8%B1%DB%8C-%D8%A7%DB%8C%D9%BE%DA%A9/"
            ],
            "prior": "بدون شماره",
            "found": {"phones": phones},
            "new_phones_vs_baseline": fresh,
            "note": "تطبیق با «پخش سراسری ایپک» قطعی نیست → زرد.",
        }
    )

    # 7) آرین پخش پیشرو — arianpakhsh.com may be different entity; keep yellow
    phones = ["09129243421"]
    fresh = count_new_phones(phones)
    for ri in find_rows(ws, "آرین پخش") + find_rows(ws, "ارین پخش"):
        set_phones(ws.cell(ri, 3), phones, uncertain=True)
    stats["new_phones"] += len(fresh)
    stats["new_phone_values"].extend(fresh)
    cases.append(
        {
            "case": "آرین پخش پیشرو",
            "queries": ["آرین پخش پیشرو تلفن", "arianpakhsh.com"],
            "sources": ["https://arianpakhsh.com/", "https://iran-gazette.com/c/4hgqsd5ebukn/%D8%A2%D8%B1%DB%8C%D9%86-%D9%BE%D8%AE%D8%B4-%D9%BE%DB%8C%D8%B4%D8%B1%D9%88"],
            "prior": "بدون شماره",
            "found": {"phones": phones},
            "new_phones_vs_baseline": fresh,
            "note": "سایت آرین پخش (زنجیره‌ای) ممکن است غیر از «آرین پخش پیشرو» گلرنگ باشد → زرد.",
        }
    )

    # 8) آسان پخش سلوا — possible تبریز آسان پخش; uncertain
    phones = ["04135247212", "09147364280"]
    fresh = count_new_phones(phones)
    for ri in find_rows(ws, "آسان پخش") + find_rows(ws, "اسان پخش"):
        set_phones(ws.cell(ri, 3), phones, uncertain=True)
        set_person(ws.cell(ri, 2), "امیر رجایی — مدیریت/سرنخ مشاغل‌برتر (تطبیق نام مبهم)", uncertain=True)
    if fresh:
        stats["new_phones"] += len(fresh)
        stats["new_phone_values"].extend(fresh)
        stats["new_people"] += 1
        stats["new_people_values"].append("امیر رجایی @ آسان پخش؟")
    cases.append(
        {
            "case": "آسان پخش سلوا",
            "queries": ["آسان پخش سلوا تلفن", "آسان پخش تبریز"],
            "sources": [
                "https://www.mbartar.ir/ad=asanpakhsh",
                "https://irangaah.ir/business/stores/tabriz/asan-pakhsh",
            ],
            "prior": "بدون شماره",
            "found": {"phones": phones, "people": ["امیر رجایی"]},
            "new_phones_vs_baseline": fresh,
            "note": "«سلوا» در منبع پیدا نشد؛ ممکن است شرکت دیگری باشد → زرد/مبهم.",
        }
    )

    # 9) بهستان پخش empty row — confirm existing HQ (not new phone)
    for ri in find_rows(ws, "بهستان پخش"):
        if not ws.cell(ri, 3).value:
            set_phones(ws.cell(ri, 3), ["02184054000"])
            set_person(ws.cell(ri, 2), "محمدهادی کمالی — مدیرعامل (تأیید سایت رسمی)", uncertain=False)
            stats["confirmed_prior"] += 1
    cases.append(
        {
            "case": "بهستان پخش",
            "queries": ["بهستان پخش تماس", "behestanpakhsh.com contact"],
            "sources": [
                "https://behestanpakhsh.com/pages/contact/",
                "https://behestanpakhsh.com/pages/team/",
            ],
            "prior": "شرکت بهستان پخش قبلاً 84054000 + کمالی دارد؛ ردیف تکراری خالی",
            "found": {"confirmed_phone": "02184054000", "confirmed_person": "محمدهادی کمالی"},
            "new_phones_vs_baseline": [],
            "note": "انتقال/تأیید شمارهٔ موجود — شمارهٔ تازه نیست.",
        }
    )

    # 10) بانی چاو تهران empty — confirm 61370 (prefix only → not new)
    for ri in find_rows(ws, "بانی چاو"):
        if not ws.cell(ri, 3).value:
            set_phones(ws.cell(ri, 3), ["02161370"])
            stats["confirmed_prior"] += 1
    # بهرام خیری current CEO clue
    append_row(
        ws,
        "شرکت بانی چاو ایرانیان",
        "بهرام خیری — مدیرعامل (سرنخ رسمیو/آخرین ترکیب)",
        [],
        uncertain=True,
    )
    stats["new_people"] += 1
    stats["new_people_values"].append("بهرام خیری @ بانی چاو")
    cases.append(
        {
            "case": "بانی چاو / پورسلطانی",
            "queries": ["علی اکبر پور سلطانی بانی چاو", "بهرام خیری بانی چاو مدیرعامل"],
            "sources": [
                "https://www.bonnychow.ir/registration",
                "https://rasmio.com/company/10861346120/%D8%A8%D8%A7%D9%86%DB%8C%20%DA%86%D8%A7%D9%88%20%D8%A7%DB%8C%D8%B1%D8%A7%D9%86%DB%8C%D8%A7%D9%86/",
            ],
            "prior": "پورسلطانی + 61370 در مبنا",
            "found": {"people": ["بهرام خیری"], "confirmed": ["02161370"]},
            "new_phones_vs_baseline": [],
        }
    )

    # --- Job moves ---
    # ابول‌نژادیان: پدیده پایدار → ماریناسان
    append_row(
        ws,
        "صنایع سلولزی ماریناسان",
        "مهران ابول نژادیان — مدیرعامل (جابه‌جایی از پخش پدیده پایدار)",
        ["02142354000", "09918546788"],
        uncertain=True,
    )
    move_phones = count_new_phones(["09918546788"])  # 42354000 exists
    stats["new_phones"] += len(move_phones)
    stats["new_phone_values"].extend(move_phones)
    stats["job_moves"] += 1
    stats["moves"].append("مهران ابول نژادیان: پدیده پایدار → ماریناسان")
    stats["new_people"] += 1
    stats["new_people_values"].append("مهران ابول نژادیان @ ماریناسان")
    cases.append(
        {
            "case": "جابه‌جایی ابول نژادیان",
            "queries": ["ابول نژادیان ماریناسان مدیرعامل", "پخش پدیده پایدار صادق بیان"],
            "sources": [
                "https://t.me/s/golrangfamily?q=%23%D9%BE%D8%AE%D8%B4_%D9%BE%D8%AF%DB%8C%D8%AF%D9%87_%D9%BE%D8%A7%DB%8C%D8%AF%D8%A7%D8%B1",
                "https://2kilopaper.ir/directory-cls_inds/listing/factories/tissue-paper/%d8%b5%d9%86%d8%a7%db%8c%d8%b9-%d8%b3%d9%84%d9%88%d9%84%d8%b2%db%8c-%d9%85%d8%a7%d8%b1%db%8c%d9%86%d8%a7%d8%b3%d8%a7%d9%86/",
                "https://www.iahci.ir/%D8%A7%D8%B9%D8%B6%D8%A7",
            ],
            "prior": "در مبنا شخص ثبت نشده بود",
            "found": {
                "move": "پدیده پایدار → ماریناسان",
                "phones": ["02142354000", "09918546788"],
            },
            "new_phones_vs_baseline": move_phones,
        }
    )

    # مهدی بهمن آبادی: بانی چاو → DOGHAZAL / غزال
    append_row(
        ws,
        "چای دوغزال / DOGHAZAL TEA",
        "مهدی بهمن آبادی — Managing Director (جابه‌جایی از بانی چاو؛ لینکدین Aug 2025)",
        [],
        uncertain=True,
    )
    stats["job_moves"] += 1
    stats["moves"].append("مهدی بهمن آبادی: بانی چاو → DOGHAZAL TEA")
    stats["new_people"] += 1
    stats["new_people_values"].append("مهدی بهمن آبادی @ DOGHAZAL")
    cases.append(
        {
            "case": "جابه‌جایی مهدی بهمن آبادی",
            "queries": ["مهدی بهمن آبادی بانی چاو", "Mehdi Bahmanabadi DOGHAZAL"],
            "sources": [
                "https://linkedin.com/in/mehdi-bahmanabadi-b1b2a479",
                "https://rasmio.com/person/467406A/%D9%85%D9%87%D8%AF%DB%8C%20%D8%A8%D9%87%D9%85%D9%86%20%D8%A2%D8%A8%D8%A7%D8%AF%DB%8C/",
            ],
            "prior": "در مبنا نبود",
            "found": {"move": "بانی چاو → DOGHAZAL TEA / عضو هیئت مدیره غزال پخش"},
            "new_phones_vs_baseline": [],
        }
    )

    # سهراب کارگر — confirm still دراژه / APIR chair (no phone)
    stats["confirmed_prior"] += 1
    cases.append(
        {
            "case": "سهراب کارگر",
            "queries": ["سهراب کارگر دراژه مدیرعامل", "apir هیئت مدیره"],
            "sources": [
                "https://www.apir.ir/pages/board-of-directors",
                "https://rasmio.com/company/10102995715/%D8%A8%D8%A7%D8%B2%D8%B1%DA%AF%D8%A7%D9%86%DB%8C%20%D8%AF%D8%B1%D8%A7%DA%98%D9%87/",
            ],
            "prior": "در مبنا بدون موبایل شخصی",
            "found": {"confirmed": "همچنان مدیرعامل دراژه و رئیس هیئت مدیره انجمن"},
            "new_phones_vs_baseline": [],
            "note": "شمارهٔ شخصی یافت نشد.",
        }
    )
    stats["no_result"] += 1  # no personal mobile for worker

    # چند پروندهٔ افراد انجمن بدون موبایل شخصی
    for name in ["مجید جعفر نژاد", "محسن عبداله زاده", "آرین نژاد"]:
        cases.append(
            {
                "case": name,
                "queries": [f"{name} موبایل", f"{name} مدیرعامل پخش"],
                "sources": ["https://www.apir.ir/pages/board-of-directors"],
                "prior": "نام در مبنا بدون موبایل کاری شخصی",
                "found": {},
                "new_phones_vs_baseline": [],
                "note": "موبایل شخصی عمومی یافت نشد؛ فقط سمت سازمانی تأیید/در دسترس است.",
            }
        )
        stats["no_result"] += 1

    # access barriers
    cases.append(
        {
            "case": "مانع دسترسی",
            "queries": ["t.me/estekhdam_tehran/93212 محتوای کامل آگهی پگاه"],
            "sources": ["https://t.me/s/estekhdam_tehran/93212"],
            "prior": "—",
            "found": {},
            "note": "صفحهٔ پیش‌نمایش تلگرام آگهی هدف را در اسنیپت نشان نداد؛ پژوهش با سایت پگاه ادامه یافت.",
            "access_barrier": True,
        }
    )
    stats["access_barrier"] += 1

    stats["cases_researched"] = len(cases)
    # dedupe counted phones
    stats["new_phone_values"] = sorted(set(stats["new_phone_values"]))
    stats["new_phones"] = len(stats["new_phone_values"])

    style_sheet(ws)
    wb.save(WORKING)
    # also keep research copy
    shutil.copy2(WORKING, LOCAL / "Distribution-Contact-Directory-FA.xlsx")

    EVID.write_text(
        json.dumps(
            {
                "generated_at": datetime.now(timezone.utc).isoformat(),
                "baseline": str(BASELINE),
                "working": str(WORKING),
                "cases": cases,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    META.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(stats, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
