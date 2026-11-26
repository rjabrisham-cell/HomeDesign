# -*- coding: utf-8 -*-
"""Import Ghasem Iran branches + APKHZ food members + Tabriz directory hits."""
from __future__ import annotations

import json
import re
import urllib.request
from pathlib import Path

OUT = Path(r"D:\HomeDesign\research\distribution-contacts\batches\person_id_hits.jsonl")
RAW = Path(r"D:\HomeDesign\research\distribution-contacts\raw")
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
    if len(d) == 7:  # Ahvaz local without area — skip ambiguous
        return None
    return d if d.startswith("0") and 10 <= len(d) <= 11 else None


def ahvaz_land(local: str) -> str | None:
    d = re.sub(r"\D", "", str(local).translate(DIGITS))
    if not d:
        return None
    if d.startswith("061"):
        return land(d)
    if len(d) == 8:
        return "061" + d
    if len(d) == 7:
        return "061" + d
    return land(d)


# Curated from search snippets + known page
GHASEM = [
    ("تهران شرق دارویی", "مهدی شاهوردی", "02166750929"),
    ("تهران غرب دارویی", "کاظم قربانی", "02144542481"),
    ("تهران غذایی", "سپهر سرمدی", "02166708542"),
    ("همدان دارویی", "سید میثم حسینیان", "08132542690"),
    ("همدان غذایی", "حسین فاضل", "08132545290"),
    ("بندرعباس/کرمان دارویی", "حبیب‌اله میرزائی", "03432610995"),
    ("بندرعباس/کرمان غذایی", "مهدی رادی", "03432610995"),
    ("ساری دارویی", "ایمان فامیلیان", "01132383764"),
    ("ساری غذایی", "محمد مراد", "01132383760"),
    ("رشت دارویی", "حسین غفاری", "01331883262"),
    ("رشت غذایی", "بهزاد پورشریف", "01333883255"),
    ("گرگان دارویی", "عبدالخالق تازیکه", "01734516351"),
    ("گرگان غذایی", "حمید سهرابی", "01734516358"),
    ("یاسوج/شیراز دارویی", "محمود کسرایی", "07137202449"),
    ("یاسوج/شیراز غذایی", "عباس سرکوبی", "07137204014"),
    ("سنندج/کرمانشاه دارویی", "محمد شباهنگی", "08334271509"),
    ("قم دارویی", "سیدعلی میرقیصری", "02536643437"),
    ("قم غذایی", "مهدی جانجانی", "02536644674"),
    ("زاهدان دارویی", "حمید سنچولی", "05433522354"),
    ("زاهدان غذایی", "محمدرضا اژدری", "05433506667"),
    ("اهواز غذایی", "سیدحسن حسینی", "06133906442"),
    ("اهواز دارویی", "هادی فاضلی‌نیا", None),
]

GHASEM_HQ = [
    ("قاسم ایران", "غلامعباس درفشی — مدیرعامل", "02144543200"),
    ("قاسم ایران", "مانی احراری — معاونت بازرگانی و فروش دارویی", "02144543147"),
    ("قاسم ایران", "محمدعلی نظرپور — مدیر برنامه‌ریزی و توسعه", "02144831116"),
    ("قاسم ایران", "ابوالفضل پلنگی — مدیر اداری و منابع انسانی", "02144543115"),
    ("قاسم ایران", "هادی قلعه‌نوی — مدیر امور مالی", "02144543141"),
    ("قاسم ایران", "مجید جهادی — مدیر فناوری اطلاعات", "02144831122"),
    ("قاسم ایران", "فرهاد امامی — مدیر لجستیک و انبارها", "02144543179"),
    ("قاسم ایران", "اعظم نافع — مدیر بازرگانی تامین کالای غیر مینو", "02144543187"),
    ("قاسم ایران", "مهناز پیرزاده‌نائینی — مدیر بازرگانی دارویی", "02144831426"),
    ("قاسم ایران", "سپهر سرمدی — مدیر فروش منطقه ۱ غذایی", "02144544946"),
    ("قاسم ایران", "محرم‌علی زیبایی — مدیر فروش منطقه ۲ غذایی", "02144549522"),
    ("قاسم ایران", "فرشید جعفری — مدیر فروش منطقه ۳ غذایی", "02144543198"),
    ("قاسم ایران", "امید مرادمند — مدیر فروش منطقه ۴ غذایی", "02144543155"),
    ("قاسم ایران", "بابک خسروانی — مدیر فروش منطقه ۱ دارویی", "02144543144"),
    ("قاسم ایران", "سید مجید آقائی — مدیر فروش منطقه ۲ دارویی", "02144543132"),
    ("قاسم ایران", "پریسا شریفی — مسئول فنی دارویی", "02144543174"),
]

# APKHZ food: company, manager, local phone (Ahvaz 061)
APKHZ = [
    ("بهفران گستر جنوب — خوزستان", "سبزیوند", "35575380"),
    ("آریاپارس دمس — خوزستان", "فرهادی", "35573044"),
    ("ارتا پخش شیرازدینا — خوزستان", "پورصدامی", "33906717"),
    ("الوند یاران سحر — خوزستان", "یاسر حردان", "33784224"),
    ("امید پخش پایدار — خوزستان", "میلاد امیدبخش", "33908267"),
    ("به پخش — خوزستان", "دغاغله", "33907010"),
    ("بهارپخش فراگیر — خوزستان", "مجدم", "35561979"),
    ("پارس خزر — خوزستان", "حیدری", "34492539"),
    ("پخش ایران (شیرین عسل) — خوزستان", "اسیوند", "32284437"),
    ("پخش پگاه (دلپذیر) — خوزستان", "داربر", "33774321"),
    ("پدیده ایده پردازان سرافراز — خوزستان", "سرافراز", "35575637"),
    ("پدیده پخش خوزستان", "امیر کیوان کتانی", "35575717"),
    ("پرنیان طعام جنوب — خوزستان", "احمد بهبهانی", "32218450"),
    ("پرنیان مهیار — خوزستان", "نقدی", "34486373"),
    ("پارس طلوع رایکا — خوزستان", "مالک", "35573294"),
    ("حیات پخش خوزستان", "رضا کتانی", "35575717"),
    ("داداش برادر (شونیز) — خوزستان", "کریم‌پور", "32270037"),
    ("زرین پخش بلک — خوزستان", "محمدعلی بلک", "35575548"),
    ("سافروشام خوزستان گستر", "ناصری", "35575733"),
    ("صفا پخش — خوزستان", "الاسوند", "35573470"),
    ("طلوع پخش آفتاب — خوزستان", "اریایی", "34427951"),
    ("گلرنگ پخش — خوزستان", "موسوی", "32288407"),
    ("گلستان — خوزستان", "محسنی", "34470273"),
    ("مدبران کوشای ایرانیان — خوزستان", "مزرعه", "35566478"),
    ("مطلوب تجارت ایرانیان — خوزستان", "سعیدپور تراب", "35575483"),
    ("میهن — خوزستان", "حیدری", "34427016"),
    ("نینا — خوزستان", "آبسالان", "33796751"),
    ("یک و یک — خوزستان", "حسینی", "33787836"),
    ("کاسپین — خوزستان", "پورصدامی", "06135547070"),
]

TABRIZ = [
    ("الیت دارو — تبریز", "مخاطب شعبه", "09143119100", "04136377037"),
    ("یاقوت سلامت — تبریز", "مخاطب شعبه", "09145884676", "04136374423"),
    ("مشاء طب — تبریز", "مخاطب شعبه", "09146301380", "04136373016"),
    ("دایا دارو — تبریز", "مخاطب شعبه", "09426002551", None),
]


def scrape_ghasem() -> list[dict]:
    url = "https://www.ghasemiran.org/fa/agency.html"
    hits = []
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=25) as r:
            html = r.read().decode("utf-8", "ignore")
    except Exception as e:
        print("ghasem fetch fail", e)
        return hits
    text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
    # list items: name - role city
    for m in re.finditer(
        r"<li[^>]*>\s*([^<\-]{3,60})\s*[-–]\s*([^<]{5,80})",
        html,
        re.I,
    ):
        pe = re.sub(r"\s+", " ", m.group(1)).strip()
        city = re.sub(r"\s+", " ", m.group(2)).strip()
        if not pe or pe.startswith("تلفن"):
            continue
        # find phone near this block
        chunk = html[m.start() : m.start() + 400]
        tel = None
        tm = re.search(r"تلفن:\s*([0-9\-۰-۹]+)", chunk)
        if tm:
            tel = land(tm.group(1))
        hits.append(
            {
                "company": f"قاسم ایران — {city}",
                "person": f"{pe} — مدیر شعبه (ghasemiran)",
                "mobile": None,
                "land": tel,
                "src": "ghasem_scrape",
            }
        )
    (RAW / "ghasem_agencies.json").write_text(
        json.dumps(hits, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return hits


def main() -> None:
    hits: list[dict] = []
    for city, pe, tel in GHASEM:
        hits.append(
            {
                "company": f"قاسم ایران — {city}",
                "person": f"{pe} — مدیر شعبه (ghasemiran)",
                "mobile": None,
                "land": land(tel) if tel else None,
                "src": "ghasem_curated",
            }
        )
    for co, pe, tel in GHASEM_HQ:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (ghasemiran managers)",
                "mobile": None,
                "land": land(tel),
                "src": "ghasem_hq",
            }
        )
    for co, pe, tel in APKHZ:
        hits.append(
            {
                "company": co,
                "person": f"{pe} — مدیر (apkhz غذایی)",
                "mobile": None,
                "land": ahvaz_land(tel),
                "src": "apkhz_food",
            }
        )
    for co, pe, mob, tel in TABRIZ:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (tebyan تبریز)",
                "mobile": to_mobile(mob),
                "land": land(tel) if tel else None,
                "src": "tabriz_dir",
            }
        )
    hits.extend(scrape_ghasem())

    existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()]
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = wm = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            if not h.get("person"):
                continue
            key = (h.get("company"), h.get("person"), h.get("mobile") or "")
            if key in seen:
                continue
            seen.add(key)
            if h.get("mobile"):
                wm += 1
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    print(json.dumps({"batch": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
