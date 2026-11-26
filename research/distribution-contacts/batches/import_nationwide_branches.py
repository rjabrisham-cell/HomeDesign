# -*- coding: utf-8 -*-
"""Import nationwide branch managers: Nokhbegan, Nafas, Karen, Behestan extras."""
from __future__ import annotations

import json
import re
import urllib.request
from pathlib import Path

ROOT = Path(r"D:\HomeDesign")
OUT = ROOT / "research/distribution-contacts/batches/person_id_hits.jsonl"
RAW = ROOT / "research/distribution-contacts/raw"
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


def land(v: str) -> str | None:
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if not d or d.startswith("09") or (len(d) == 10 and d.startswith("9")):
        return None
    if d.startswith("98"):
        d = "0" + d[2:]
    if not d.startswith("0") and len(d) == 10:
        d = "0" + d
    return d if d.startswith("0") and len(d) >= 10 else None


NOKHBEGAN = [
    ("اصفهان", "سید محمدحسین مصلحی", "03145838271"),
    ("تهران غرب", "محمدقاسم طیبی‌زاد", "02187792000"),
    ("شیراز", "محمدرضا حسن‌زاده", "07132426522"),
    ("شهرکرد", "سید سجاد صالحی", "03833336190"),
    ("البرز", "محسن مهرپرور", ""),
    ("اردبیل", "اسداله شارقی", ""),
    ("تبریز", "نادر وظیفه‌مهربانی", "04136378928"),
    ("ارومیه", "اصغر تقی‌نژاد", "04432722598"),
    ("رشت", "معصومه خباز", "01333691053"),
    ("گرگان", "حمید سوراوند", "01734534291"),
    ("همدان", "محمد محمدی", ""),
    ("ایلام", "یونس سیاهی", "08432234368"),
    ("کرمان", "فرهاد سلمانی", "02187792000"),
    ("کرمانشاه", "حمید باتمانی", "08334281037"),
    ("مشهد", "احسان پیله‌ور", "05136511061"),
    ("اهواز", "مصطفی زاروی", "02187792000"),
    ("خرم‌آباد", "رضا جمالی", "06633435106"),
    ("ساری", "محمدزمان فرجی", "02187792000"),
    ("قزوین", "روشن انصاری امین‌جان", "02832376617"),
    ("قم", "حسین حسینی", "02536641137"),
    ("یزد", "مجتبی ملانوری شمسی", "03535262884"),
]

NAFAS = [
    ("کرج/البرز", "مسعود سلطانی", "02122581404"),
    ("تهران شرق", "مجتبی جلالی", "02126373548"),
    ("تهران غرب", "عباس اورنگی", "02144939285"),
    ("مشهد", "سعید احمدی", ""),
    ("تبریز", "اسمعیل سربازنژاد", "04136370744"),
    ("شیراز", "ایوب سلطانی", "07137732032"),
]

# Karen curated from search + known pages (city, person, mobile, land)
KAREN_CURATED = [
    ("تهران", "رحیمی", "09195009334", "02188099960"),
    ("تهران شرق", "حسن‌زاده", "09120766053", "02188099960"),
    ("خراسان رضوی", "مجتبی حسینی", "09158049099", "05136662021"),
    ("کرمان", "کاظمی‌پور", "09133415987", "03433351000"),
    ("مازندران", "شاکری", "09127347838", "01133286501"),
    ("فارس", "سرگل‌زائی", "09173401355", "07137730016"),
    ("فارس", "سجادیان — سرپرست فروش", "09177009880", "07137730016"),
    ("البرز", "رضا حق‌رو", "09166050306", ""),
    ("قم", "بختیاری", "09122519492", ""),
    ("همدان", "اعتمادی‌طلب", "09183129645", ""),
]

BEHESTAN_MORE = [
    ("قزوین", "امیرحسین وهاب‌پور — سرپرست فروش", "02834318897"),
]


def scrape_karen() -> list[dict]:
    hits = []
    for i in range(1, 50):
        url = f"https://karendistribution.com/branch/show/{i}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as r:
                html = r.read().decode("utf-8", "ignore")
        except Exception:
            continue
        title_m = re.search(r"<h1[^>]*>(.*?)</h1>", html, re.I | re.S)
        city = re.sub(r"<[^>]+>", "", title_m.group(1)).strip() if title_m else f"شعبه {i}"
        pe = None
        m = re.search(
            r"نام و نام خانوادگی</t[dh]>\s*<t[dh][^>]*>\s*([^<]{3,80})",
            html,
            re.I | re.S,
        )
        if m:
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
        if not pe:
            m = re.search(r"(آقای|خانم)\s+([^<]{2,40})", html)
            if m:
                pe = (m.group(1) + " " + m.group(2)).strip()
        mobs = []
        for mm in re.finditer(r"(?<!\d)(0?9\d{9})(?!\d)", html.translate(DIGITS)):
            mb = to_mobile(mm.group(1))
            if mb and mb not in mobs:
                mobs.append(mb)
        lands = []
        for mm in re.finditer(r"(?<!\d)(0\d{2}\d{7,8})(?!\d)", html.translate(DIGITS)):
            d = mm.group(1)
            if d.startswith("09"):
                continue
            if d not in lands:
                lands.append(d)
        if pe or mobs:
            hits.append(
                {
                    "company": f"پخش سراسری کارن — شعبه {city}",
                    "person": f"{pe or 'مخاطب شعبه'} — مدیر/مخاطب شعبه (karendistribution)",
                    "mobile": mobs[0] if mobs else None,
                    "land": lands[0] if lands else None,
                    "src": "karen_scrape",
                    "extra_mobiles": mobs[1:],
                }
            )
    (RAW / "karen_branches.json").write_text(
        json.dumps(hits, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return hits


def main() -> None:
    hits: list[dict] = []
    for city, pe, tel in NOKHBEGAN:
        hits.append(
            {
                "company": f"داروگستر نخبگان — شعبه {city}",
                "person": f"{pe} — مدیر شعبه (nokhbegandc.com)",
                "mobile": None,
                "land": land(tel),
                "src": "nokhbegan_branches",
            }
        )
    for city, pe, tel in NAFAS:
        hits.append(
            {
                "company": f"نیک فردای سالم (نفس) — شعبه {city}",
                "person": f"{pe} — رئیس شعبه (nafaspharma.com)",
                "mobile": None,
                "land": land(tel),
                "src": "nafas_branches",
            }
        )
    for city, pe, mob, tel in KAREN_CURATED:
        hits.append(
            {
                "company": f"پخش سراسری کارن — شعبه {city}",
                "person": f"{pe} — مدیر/سرپرست شعبه (karendistribution)",
                "mobile": to_mobile(mob),
                "land": land(tel) if tel else None,
                "src": "karen_curated",
            }
        )
    for city, pe, tel in BEHESTAN_MORE:
        hits.append(
            {
                "company": f"بهستان پخش — شعبه {city}",
                "person": f"{pe} (behestanpakhsh.com)",
                "mobile": None,
                "land": land(tel),
                "src": "behestan_qazvin",
            }
        )

    scraped = scrape_karen()
    hits.extend(scraped)
    # explode extra mobiles as additional rows
    extra = []
    for h in scraped:
        for m in h.get("extra_mobiles") or []:
            extra.append(
                {
                    "company": h["company"],
                    "person": h["person"] + " — تماس اضافی",
                    "mobile": m,
                    "land": h.get("land"),
                    "src": "karen_extra_mobile",
                }
            )
    hits.extend(extra)

    existing = []
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if line.strip():
                existing.append(json.loads(line))
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = 0
    with_mob = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            if not h.get("person"):
                continue
            key = (h.get("company"), h.get("person"), h.get("mobile") or "")
            if key in seen:
                continue
            seen.add(key)
            if h.get("mobile"):
                with_mob += 1
            out = {
                "company": h["company"],
                "person": h["person"],
                "mobile": h.get("mobile"),
                "land": h.get("land"),
                "src": h.get("src"),
            }
            f.write(json.dumps(out, ensure_ascii=False) + "\n")
            added += 1
    print(
        json.dumps(
            {
                "batch": len(hits),
                "appended": added,
                "with_mobile": with_mob,
                "karen_scraped": len(scraped),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
