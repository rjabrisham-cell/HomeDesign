# -*- coding: utf-8 -*-
"""Import Exir branches (with mobiles), Qazvin FDO pairs, Behestan HQ+branches."""
from __future__ import annotations

import json
import re
from pathlib import Path

OUT = Path(r"D:\HomeDesign\research\distribution-contacts\batches\person_id_hits.jsonl")
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def to_mobile(v: str | None) -> str | None:
    if not v:
        return None
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    if len(d) == 11 and d.startswith("09"):
        return d
    return None


def land(v: str | None) -> str | None:
    if not v:
        return None
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if not d or d.startswith("09") or (len(d) == 10 and d.startswith("9")):
        return None
    if d.startswith("98"):
        d = "0" + d[2:]
    if not d.startswith("0") and len(d) == 10:
        d = "0" + d
    if len(d) == 8:
        d = "021" + d
    return d if d.startswith("0") and 10 <= len(d) <= 11 else None


# Exir: (city, person, mobile, land)
EXIR = [
    ("تبریز", "مرتضی جسور", "09144044409", "04133147022"),
    ("ارومیه", "وحید پاشایی‌لور", "09143495426", "04433745130"),
    ("اصفهان", "اصغر پناهنده", "09134208089", "03135723724"),
    ("تهران غرب", "یعقوب بدلیان", "09123682311", "02165607645"),
    ("تهران شرق", "حسین کسائیان", "09126070015", "02188545943"),
    ("بیرجند", "میثم صباغی", "09153611040", "05632255052"),
    ("مشهد", "محمد سلطانی", None, "05136220166"),
    ("اهواز", "میثم حمزه‌لو", None, "0613297116"),
    ("زنجان", "علی قرابگلی", None, "02432221243"),
    ("زاهدان", "مهدی آرام", None, "05433592297"),
    ("شیراز", "علی نجفی", None, "07138408387"),
    ("قم", "محمد وزینی", None, "02536642220"),
    ("کرمان", "حسنعلی بلالی", "09138909010", "03432752944"),
    ("کرمانشاه", "اسماعیل علم‌خواه", "09128953274", "08338353011"),
    ("گرگان", "کریم کوثری", "09112718066", "01732170934"),
    ("رشت", "ایمان قربان‌پور", "09118274048", "01331883183"),
]

EXIR_HQ = [
    ("پخش اکسیر", "محمدمهدی رمضانی‌پور — مدیرعامل (exirpd)", None),
    ("پخش اکسیر", "حمیدرضا مبرهن — معاون مالی اقتصادی (exirpd)", None),
    ("پخش اکسیر", "حسین قمی‌زاده — معاون اجرایی (exirpd)", None),
    ("پخش اکسیر", "محمد میرزایی — معاون بازرگانی (exirpd)", None),
    ("پخش اکسیر", "امیر علائی — معاون فروش (exirpd)", None),
    ("پخش اکسیر", "یوسف زنگیدن — مشاور لجستیک (exirpd)", None),
    ("پخش اکسیر", "ابراهیم طاهری‌مهر — مدیر حراست (exirpd)", None),
    ("پخش اکسیر", "سحر نوری — مسئول فنی (exirpd)", None),
    ("پخش اکسیر", "محمد الماسی — مدیر فروش منطقه (exirpd)", None),
    ("پخش اکسیر", "سعید توانا — مدیر فروش منطقه (exirpd)", None),
    ("پخش اکسیر", "امین ولدخانی — مدیر فروش منطقه (exirpd)", None),
    ("پخش اکسیر", "محسن کریمی — مدیر خرید تجهیزات (exirpd)", None),
    ("پخش اکسیر", "صمد آقالری — مدیر مالی (exirpd)", None),
    ("پخش اکسیر", "کمال طیبی‌زاد — مدیر انبار مرکزی (exirpd)", None),
]

# Qazvin FDO: (company, manager, technical)
QAZVIN = [
    ("داروپخش — قزوین", "میار صنمی — مدیر/فروش", "بهناز نیاقی — مسئول فنی"),
    ("پخش فردوس — قزوین", "دباغ — مدیر", "لیلا طیبی جزایری — مسئول فنی"),
    ("پخش هجرت — قزوین", "سلیمانی — مدیر", "فرامرز کارآگاه — مسئول فنی"),
    ("داروسازان التیام — قزوین", "زارعی — مدیر", "صدیقه کلانتری — مسئول فنی"),
    ("آدرا طب — قزوین", "شفاهی — مدیر", "ثمین نظری — مسئول فنی"),
    ("داروگستر نخبگان — قزوین", "انصاری — مدیر", "معصومه رفیع‌خواه — مسئول فنی"),
    ("شفا آراد — قزوین", "امینی — مدیر", "سایه ابراهیمی — مسئول فنی"),
    ("قاسم ایران — قزوین", "نعمتی — مدیر", "مرتضی حسینی — مسئول فنی"),
    ("سینا پخش ژن — قزوین", "حسینی — مدیر فروش", "ناصر بابا خداوردی — مسئول فنی"),
    ("بهستان پخش — قزوین", "کامیاب محمدی — مدیر", "شایان وطن‌دوست — مسئول فنی"),
    ("میلان پارس فارمد — قزوین", "کوشان — مدیر فروش", "یاسمن پوربابا — مسئول فنی"),
    ("سایه سمن — قزوین", "رمضانی — مدیر", "سینوش کوشان — مسئول فنی"),
    ("سها سروش نوین آسمان — قزوین", "حسین‌لو — مدیر", "زهرا اصغری‌ها — مسئول فنی"),
    ("دهکدک — قزوین", "شعبانی — مدیر", None),
]

BEHESTAN_MORE = [
    ("مازندران", "صابر اسمعیل‌زاده — مدیر شعبه", "01133347530"),
    ("مازندران", "مهدی مهدی‌زاده — سرپرست فروش", "01133347530"),
    ("همدان", "ابراهیم مرادی — مدیر شعبه", "08132569237"),
    ("همدان", "امیر خدابنده — سرپرست فروش", "08132569237"),
    ("کرمانشاه", "بیژن چقازردی — مدیر مرکز", "02184054994"),
    ("کرمانشاه", "فرید فیض‌نیا — سرپرست فروش", "02184054994"),
    ("یزد", "محمد صالحی‌راد — مدیر شعبه", ""),
    ("یزد", "محمدحسن غیاثی — سرپرست فروش", ""),
]

BEHESTAN_HQ = [
    ("بهستان پخش", "محمدهادی کمالی — مدیرعامل (behestan team)"),
    ("بهستان پخش", "رضا اشجعی — معاون فروش"),
    ("بهستان پخش", "رئوف اسدیان — مدیر خرید"),
    ("بهستان پخش", "بابک میثاقی — مدیر فناوری اطلاعات"),
    ("بهستان پخش", "هومن آجودانی — مدیر حسابداری فروش"),
    ("بهستان پخش", "صبا ساسانی — مدیر حسابداری مالی"),
    ("بهستان پخش", "مهدی احمدپور علمداری — مدیر فنی و پشتیبانی"),
    ("بهستان پخش", "مهیار قربانی — مدیر برنامه‌ریزی استراتژیک"),
    ("بهستان پخش", "مجتبی ابوالفتحی — مدیر حسابرسی داخلی"),
    ("بهستان پخش", "لیدا بهادر — مدیر فنی دارویی"),
    ("بهستان پخش", "حسین پرندوار — مدیر حقوقی"),
]


def main() -> None:
    hits: list[dict] = []
    for city, pe, mob, tel in EXIR:
        hits.append(
            {
                "company": f"پخش اکسیر — مرکز {city}",
                "person": f"{pe} — مدیر مرکز (exirpd.com)",
                "mobile": to_mobile(mob),
                "land": land(tel),
                "src": "exir_branches",
            }
        )
    for co, pe, tel in EXIR_HQ:
        hits.append(
            {
                "company": co,
                "person": pe,
                "mobile": None,
                "land": land(tel) if tel else None,
                "src": "exir_hq",
            }
        )
    for co, mgr, tech in QAZVIN:
        hits.append(
            {
                "company": co,
                "person": f"{mgr} (FDO قزوین)",
                "mobile": None,
                "land": None,
                "src": "qazvin_fdo",
            }
        )
        if tech:
            hits.append(
                {
                    "company": co,
                    "person": f"{tech} (FDO قزوین)",
                    "mobile": None,
                    "land": None,
                    "src": "qazvin_fdo",
                }
            )
    for city, pe, tel in BEHESTAN_MORE:
        hits.append(
            {
                "company": f"بهستان پخش — {city}",
                "person": f"{pe} (behestanpakhsh)",
                "mobile": None,
                "land": land(tel) if tel else None,
                "src": "behestan_h2",
            }
        )
    for co, pe in BEHESTAN_HQ:
        hits.append(
            {
                "company": co,
                "person": pe,
                "mobile": None,
                "land": None,
                "src": "behestan_hq",
            }
        )

    existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()]
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = 0
    with_mob = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            key = (h.get("company"), h.get("person"), h.get("mobile") or "")
            if key in seen:
                continue
            seen.add(key)
            if h.get("mobile"):
                with_mob += 1
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    print(json.dumps({"batch": len(hits), "appended": added, "with_mobile": with_mob}, ensure_ascii=False))


if __name__ == "__main__":
    main()
