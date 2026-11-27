# -*- coding: utf-8 -*-
"""Import Biodep reps, Syndi CEOs, richer Fars PDF parse, association leftovers."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "batches" / "person_id_hits.jsonl"
RAW = ROOT / "raw"

DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def to_mobile(v):
    if not v:
        return None
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    return d if len(d) == 11 and d.startswith("09") else None


def land(v):
    if not v:
        return None
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if not d or d.startswith("09"):
        return None
    if not d.startswith("0") and len(d) in (10, 11):
        d = "0" + d
    return d if d.startswith("0") and 10 <= len(d) <= 11 else None


def append_hits(hits: list[dict]) -> tuple[int, int]:
    existing = []
    if OUT.exists():
        existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()]
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = wm = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            key = (h["company"], h["person"], h.get("mobile") or "")
            if key in seen:
                continue
            seen.add(key)
            if h.get("mobile"):
                wm += 1
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    return added, wm


BIODEP = [
    ("پویا پخش اسپادانا — اصفهان", "دکتر قاسمی نژاد", "03135722101"),
    ("مهر نیکا دارو — تهران", "دکتر فرزاد حیدری", "02126219176"),
    ("پرند پیروز مهاباد — آذربایجان غربی", "دکتر پیمان بامداد", "04442243020"),
    ("فرساد ماکیان دارو — آذربایجان شرقی", "دکتر داوطلب", "04137454794"),
    ("نیک پخش رسا — آذربایجان شرقی", "دکتر رسول عباسخانی", "04134244123"),
    ("البرز پارس پویش — مازندران", "دکتر رحمانزاده", "01143249134"),
    ("آبان دارو گستر — کهگیلویه و بویراحمد", "خانم دکتر مردانی", "07433354553"),
    ("پارسیان پخش آرنا — گیلان", "مهندس فرزاد علی اکبرخانی", "01333559242"),
    ("ماکیان دارو گلستان", "خانم معصومه سعیدی", "01732521755"),
    ("زرین پخش پدیده کومش — سمنان", "دکتر محمود حسنعلی زاده روشن", "02333652542"),
    ("اترک دارو زنجان", "دکتر رامین بخشی رحمتی", "02433756219"),
    ("ماهدشت داروی ایرانیان — البرز", "دکتر امیر ناصر مزینانی", "02637304387"),
    ("پرنیان طب زاگرس — چهارمحال", "دکتر سید مهداد آل علی", "03832272191"),
    ("گیتی پخش چکاوک — خوزستان", "دکتر سعید نیل زن", "06142261426"),
    ("افلاک پخش ملک — لرستان", "دکتر ملک", "06643423374"),
    ("سرو دام بیرجند — خراسان جنوبی", "مهندس علی خاکستانی شهری", "05632318156"),
    ("صحت دام یزد", "دکتر سید احمد کمال آبادی", "03537223380"),
    ("دارو برنا — فارس", "مهندس شهرام دشتی زاده", None),
    ("همت پخش بوتیا — کرمان", "دکتر حکیمه مظفری", "03432730747"),
    ("فرزام دارو پخش پارس — خراسان رضوی", "دکتر الیاس نوری", "05136779654"),
    ("به ثمین درمان کیان — قم", "خانم دکتر معصومه اکبری کیان", "02536625446"),
]

SYNDI = [
    ("پخش اکسیر", "دکتر محمد هادی کمالی", "02177687801"),
    ("پخش البرز", "دکتر عقیل آرین نژاد", "02166707727"),
    ("توزیع داروپخش", "بیژن کلانتری اشان", "02188888073"),
    ("پخش رازی", "مهندس محمدرضا اسدآبادی", "02166715712"),
    ("پخش فردوس", "دکتر مجید زجاجی", "02166035170"),
    ("قاسم ایران", "مهندس علی محمد وحید کریمی", "02144543200"),
    ("پخش هجرت", "دکتر مسعود بهرامی نوید", "02166461933"),
    ("پخش مهبان دارو", "مهندس علی باقری", "02144626506"),
]

# Richer Fars rows from PDF text
FARS = [
    ("پخش هجرت — فارس", "حسن زاده", "09178871255", "07138247775"),
    ("سایه سمن — فارس", "بایزدان", "09178200919", "07137418682"),
    ("قاسم ایران — فارس", "کسرایی", "09171128581", "07137202449"),
    ("قاسم ایران — فارس", "سرکوبی", "09173101287", None),
    ("پخش البرز — فارس", "علیزاده", "09173000114", "07191010515"),
    ("پخش رازی — فارس", "زارع", "09171182965", "07132425316"),
    ("توزیع داروپخش — فارس", "کشاورز", "09177169647", "07132268590"),
    ("بهستان پخش — فارس", "یزدان خواه", "09173031085", "02184054982"),
    ("پخش ممتاز — فارس", "نجفی", "09365826398", "07137425475"),
    ("آدوراطب — فارس", "زارعی", "09171125238", "07138221244"),
    ("پخش فردوس — فارس", "ترابیان", "09177025877", "07138424748"),
    ("محیا دارو — فارس", "حسینی راد", "09179327173", "07137742323"),
    ("پخش مکتاف — فارس", "مخاطب شعبه مکتاف", None, "07137732794"),
    ("پخش کارن — فارس", "سرگل زایی", "09128106295", "07137730016"),
    ("داروسازان التیام — فارس", "صدر", "09173110024", "07137740341"),
    ("میلاد داروی فارس", "افشار", "09177878633", "07138228080"),
    ("میلاد داروی فارس", "بیات", "09178961311", None),
    ("داروگستر پیشرو پارسیان — فارس", "مهشید فروزش", "09173142377", "07132321002"),
    ("داروگستر پیشرو پارسیان — فارس", "موسوی", "09178011803", None),
    ("شرکت دارویی اکسیر — فارس", "سلطانی", "09179195250", "07138408387"),
    ("پورا پخش — فارس", "زاهدی", None, "07132427804"),
    ("مشاء طب — فارس", "دهداری", "09171132280", None),
    ("شرکت سها هلال — فارس", "بهادری", "09177154924", "07138424219"),
    ("شرکت سها هلال — فارس", "پرتو", "09395624255", None),
    ("باریج اسانس — فارس", "قاسم رکن پور", "09177052329", "07132426360"),
    ("الیت دارو — فارس", "شریفی نژاد", "09171062755", None),
    ("سلامت پخش هستی — فارس", "سبزعلیان", "09171088556", "07137743258"),
    ("طب گران نوین — فارس", "احمدی", "09175388848", "07132267049"),
]


def assoc_leftovers() -> list[dict]:
    path = ROOT / "association_members_raw.json"
    if not path.exists():
        return []
    rows = json.loads(path.read_text(encoding="utf-8"))
    hits = []
    for r in rows:
        co = (r.get("name_raw") or "").strip()
        pe = (r.get("ceo_source") or "").strip()
        phone = r.get("phone_official") or ""
        if not co or not pe:
            continue
        # normalize phone fragments like 60912401-2 / 62937
        ld = None
        d = re.sub(r"\D", "", str(phone).translate(DIGITS))
        if d:
            if not d.startswith("0") and len(d) >= 8:
                # assume Tehran 021 if looks short
                if len(d) == 8:
                    d = "021" + d
                elif len(d) == 10:
                    d = "0" + d
            ld = land(d)
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیرعامل فهرست انجمن ملی)",
                "mobile": to_mobile(phone),
                "land": ld,
                "src": "assoc_ceo_pass",
            }
        )
    return hits


def parse_fars_txt() -> list[dict]:
    path = RAW / "fars_members_1404.txt"
    if not path.exists():
        return []
    text = path.read_text(encoding="utf-8").translate(DIGITS)
    hits = []
    # extract all 09 mobiles as company+mobile with person stub unique
    for mb in sorted(set(re.findall(r"09\d{9}", text))):
        hits.append(
            {
                "company": f"عضو انجمن پخش فارس — همراه {mb}",
                "person": f"مخاطب همراه {mb[-4:]} (فهرست فارس ۱۴۰۴)",
                "mobile": mb,
                "land": None,
                "src": "fars_pdf_all_mobs",
            }
        )
    return hits


def main():
    hits: list[dict] = []
    for co, pe, phone in BIODEP:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (نماینده استانی زیست‌درمان ماهان)",
                "mobile": None,
                "land": land(phone) if phone else None,
                "src": "biodep_full",
            }
        )
    for co, pe, phone in SYNDI:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیرعامل سندیکای دارو)",
                "mobile": None,
                "land": land(phone),
                "src": "syndi_ceo",
            }
        )
    for co, pe, mob, ld in FARS:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیر انجمن فارس)",
                "mobile": to_mobile(mob),
                "land": land(ld) if ld else None,
                "src": "fars_rich",
            }
        )
    hits.extend(parse_fars_txt())
    assoc = assoc_leftovers()
    print({"assoc_candidates": len(assoc)})
    hits.extend(assoc)
    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
