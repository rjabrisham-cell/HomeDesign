# -*- coding: utf-8 -*-
"""Fast curated volume import: Nokhbegan, QUMS, Cina, Fars mobiles, Syndi, Biodep."""
from __future__ import annotations

import json
import re
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "batches" / "person_id_hits.jsonl"
RAW = ROOT / "raw"
RAW.mkdir(parents=True, exist_ok=True)

DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}


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


NOKHBEGAN = [
    ("اصفهان", "سید محمد حسین مصلحی", "03145838271"),
    ("تهران غرب", "محمدقاسم طیبی زاد", "02187792000"),
    ("شیراز", "محمدرضا حسن زاده", "07132426522"),
    ("شهرکرد", "سید سجاد صالحی", "03833336190"),
    ("البرز", "محسن مهرپرور", None),
    ("اردبیل", "اسداله شارقی", None),
    ("تبریز", "نادر وظیفه مهربانی", "04136378928"),
    ("ارومیه", "اصغر تقی نژاد", "04432722598"),
    ("رشت", "معصومه خباز", "01333691053"),
    ("گرگان", "حمید سوراوند", "01734534291"),
    ("همدان", "محمد محمدی", None),
    ("ایلام", "یونس سیاهی", "08432234368"),
    ("کرمان", "فرهاد سلمانی", None),
    ("کرمانشاه", "حمید باتمانی", "08334281037"),
    ("مشهد", "احسان پیله ور", "05136511061"),
    ("اهواز", "مصطفی زاروی", None),
    ("لرستان", "رضا جمالی", "06633435106"),
    ("ساری", "محمد زمان فرجی", None),
    ("قزوین", "روشن انصاری", "02832376617"),
]

QUMS = [
    ("پخش فردوس — قزوین", "دباغ", "نیاقى", "02833292571"),
    ("پخش هجرت — قزوین", "خوبان", "فرامرز کارآگاه", "02833297591"),
    ("پخش التیام — قزوین", "کشاورز", "جمالی", "02833292252"),
    ("آدوراطب — قزوین", "شفاهی", "شاه تقی", "02832848155"),
    ("داروگستر نخبگان — قزوین", "انصاری", "میررحیمی", "02832376617"),
    ("قاسم ایران — قزوین", "مظفرپور", "فتحی", "02832580220"),
    ("دی دارو امید — قزوین", "نصیری", None, "02832224501"),
    ("داروگستر شایان اعتماد — قزوین", "مرادی", "سیاهکلی", "02833290831"),
    ("داروپخش — قزوین", "میار صنمی", "بهناز نیاقی", None),
    ("شفا آراد — قزوین", "امینی", "سایه ابراهیمی", None),
    ("سینا پخش ژن — قزوین", "حسینی", "ناصر بابا خداوردی", None),
    ("بهستان پخش — قزوین", "کامیاب محمدی", "شایان وطن دوست", None),
    ("میلان پارس فارمد — قزوین", "کوشان", "یاسمن پور بابا", None),
    ("سایه سمن — قزوین", "رمضانی", "سینوش کوشان", None),
    ("سها سروش نوین آسمان — قزوین", "حسین لو", "زهرا اصغری ها", None),
    ("دهکدک — قزوین", "شعبانی", None, None),
]

CINA = [
    ("غرب تهران", "حمیدرضا حبیب زاده", "02144119010", None),
    ("شرق تهران", "احد صفرخانلو", "02177248913", None),
    ("تخصصی تهران", "بیژن عباسی", "02144664308", None),
    ("مشهد", "رضا خانی", "02191074460", None),
    ("شیراز", "سیدغلامرضا علوی", "07137742956", None),
]

# Fars association PDF snippet: company, manager, mobile
FARS = [
    ("پخش هجرت — فارس", "حسن", "09178871255", "07138247775"),
    ("سایه سمن — فارس", "بایزدان", "09178200919", "07137418682"),
    ("قاسم ایران — فارس", "کسرایی", "09171128581", "07137202449"),
    ("پخش البرز — فارس", "علیزاده", "09173000114", "07191010515"),
    ("پخش رازی — فارس", "زارع", "09171182965", "07132425316"),
    ("توزیع داروپخش — فارس", "کشاورز", "09177169647", "07132268590"),
    ("بهستان پخش — فارس", "یزدان خواه", "09173031085", "02184054982"),
    ("پخش ممتاز — فارس", "نجفی", "09365826398", "07137425475"),
    ("آدوراطب — فارس", "زارعی", "09171125238", "07138221244"),
    ("پخش فردوس — فارس", "ترابیان", "09177025877", "07138424748"),
    ("محیا دارو — فارس", "حسینی", "09179327173", "07137742323"),
    ("پخش کارن — فارس", "سرگل زایی", "09128106295", "07137730016"),
    ("داروسازان التیام — فارس", "صدر", "09173110024", "07137740341"),
    ("میلاد داروی فارس", "افشار", "09177878633", "07138228080"),
    ("داروگستر پیشرو پارسیان — فارس", "مهشید موسوی", "09173142377", "07132321002"),
    ("شرکت دارویی اکسیر — فارس", "سلطانی", "09179195250", "07138408387"),
    ("پورا پخش — فارس", "زاهدی", None, "07132427804"),
    ("مشاء طب — فارس", "دهداری", "09171132280", None),
    ("شرکت سها هلال — فارس", "بهادری", "09177154924", "07138424219"),
    ("باریج اسانس — فارس", "قاسم رکن پور", "09177052329", "07132426360"),
    ("الیت دارو — فارس", "شریفی نژاد", "09171062755", None),
    ("سلامت پخش هستی — فارس", "سبزعلیان", "09171088556", "07137743258"),
    ("طب گران نوین — فارس", "احمدی", "09175388848", "07132267049"),
]

BIODEP = [
    ("مهر نیکا دارو", "دکتر فرزاد حیدری", "02126219176"),
    ("دارو برنا — فارس", "مهندس شهرام دشتی زاده", "07138013"),
    ("فرزام دارو پخش پارس — مشهد", "دکتر الیاس نوری", "05136779654"),
]

EXTRA = [
    ("یاراطب ثامن", "سیاوش مروج تربتی", None, "02144083412"),
    ("پخش هجرت", "مسعود بهرامی نوید", None, "02166461091"),
    ("داروگستر نخبگان", "بهزاد آسوده", None, "02187792000"),
    ("پخش اکسیر", "دکتر محمد هادی کمالی", None, "02177687801"),
]


def fetch_cina() -> list[dict]:
    hits = []
    try:
        req = urllib.request.Request("https://cinnadco.com/", headers=UA)
        html = urllib.request.urlopen(req, timeout=20).read().decode("utf-8", "ignore")
        (RAW / "cinnadco.html").write_text(html, encoding="utf-8")
        text = re.sub(r"<[^>]+>", "\n", html)
        text = re.sub(r"\n+", "\n", text).translate(DIGITS)
        for m in re.finditer(
            r"نام مرکز\s*:\s*([^\n]+)\s*نام مدیر\s*:\s*([^\n]+)\s*شماره تماس\s*:\s*([^\n]+)",
            text,
        ):
            city = re.sub(r"\s+", " ", m.group(1)).strip()
            pe = re.sub(r"\s+", " ", m.group(2)).strip()
            phone = re.sub(r"\s+", " ", m.group(3)).strip()
            mb = to_mobile(phone)
            ld = None if mb else land(phone if phone.startswith("0") else "0" + phone)
            hits.append(
                {
                    "company": f"سینا پخش — مرکز {city}",
                    "person": f"{pe} (مدیر مرکز سینا پخش {city})",
                    "mobile": mb,
                    "land": ld,
                    "src": "cina_scrape",
                }
            )
        print({"cina_scrape": len(hits)})
    except Exception as e:
        print({"cina_error": str(e)})
    return hits


def fetch_fars_pdf() -> list[dict]:
    hits = []
    url = "https://adfars.ir/wp-content/uploads/2025/12/%D9%84%DB%8C%D8%B3%D8%AA-%D8%B4%D8%B1%DA%A9%D8%AA%D9%87%D8%A7%DB%8C-%D9%BE%D8%AE%D8%B4-%D8%AF%D8%A7%d8%b1%d9%88%d8%a6%db%8c-%d8%b9%d8%b6%d9%88-%d8%b5%d9%86%d8%b9%d8%aa-%d9%be%d8%ae%d8%b4-%d8%a7%d8%b3%d8%aa%d8%a7%d9%86-%d9%81%d8%a7%d8%b1%d8%b3-1404.10.02.pdf"
    # try simpler known local raw first
    local = RAW / "fars_adfars.pdf"
    try:
        req = urllib.request.Request(
            "https://adfars.ir/wp-content/uploads/2025/12/"
            + "%D9%84%DB%8C%D8%B3%D8%AA-%D8%B4%D8%B1%DA%A9%D8%AA%D9%87%D8%A7%DB%8C-%D9%BE%D8%AE%D8%B4-%D8%AF%D8%A7%D8%B1%D9%88%D8%A6%DB%8C-%D8%B9%D8%B6%D9%88-%D8%B5%D9%86%D8%B9%D8%AA-%D9%BE%D8%AE%D8%B4-%D8%A7%D8%B3%D8%AA%D8%A7%D9%86-%D9%81%D8%A7%D8%B1%D8%B3-1404.10.02.pdf",
            headers=UA,
        )
        data = urllib.request.urlopen(req, timeout=30).read()
        p = RAW / "fars_members_1404.pdf"
        p.write_bytes(data)
        print({"fars_pdf_bytes": len(data)})
        # extract text via pdfminer/pypdf if available
        text = ""
        try:
            from pypdf import PdfReader

            reader = PdfReader(str(p))
            text = "\n".join((pg.extract_text() or "") for pg in reader.pages)
        except Exception:
            try:
                import pdfminer.high_level

                text = pdfminer.high_level.extract_text(str(p))
            except Exception as e:
                print({"pdf_parse_error": str(e)})
        if text:
            (RAW / "fars_members_1404.txt").write_text(text, encoding="utf-8")
            # mobiles
            for mb in set(re.findall(r"09\d{9}", text.translate(DIGITS))):
                hits.append(
                    {
                        "company": "عضو انجمن پخش فارس (PDF 1404)",
                        "person": f"مخاطب همراه {mb[-4:]} (فهرست فارس)",
                        "mobile": mb,
                        "land": None,
                        "src": "fars_pdf_mobile",
                    }
                )
    except Exception as e:
        print({"fars_pdf_error": str(e)})
        if local.exists():
            print({"using_local_fars": True})
    return hits


def fetch_syndi() -> list[dict]:
    hits = []
    try:
        url = "https://syndipharma.org/%d9%84%db%8c%d8%b3%d8%aa-%d8%a2%d8%af%d8%b1%d8%b3-%d9%88-%d8%aa%d9%84%d9%81%d9%86-%d8%aa%d9%85%d8%a7%d8%b3-%d8%b4%d8%b1%da%a9%d8%aa%d9%87%d8%a7%db%8c-%d9%be%d8%ae%d8%b4-%d8%b9%d8%b6%d9%88-%d9%88%d8%a7/"
        req = urllib.request.Request(url, headers=UA)
        html = urllib.request.urlopen(req, timeout=25).read().decode("utf-8", "ignore")
        (RAW / "syndipharma_pakhsh.html").write_text(html, encoding="utf-8")
        text = re.sub(r"<[^>]+>", "\n", html)
        text = re.sub(r"\n+", "\n", text)
        # rough: company name lines near مدیر
        # keep curated from snippet for reliability
    except Exception as e:
        print({"syndi_error": str(e)})
    return hits


def main():
    hits: list[dict] = []
    for city, pe, phone in NOKHBEGAN:
        hits.append(
            {
                "company": f"داروگستر نخبگان — شعبه {city}",
                "person": f"{pe} (مدیر شعبه نخبگان {city})",
                "mobile": None,
                "land": land(phone) if phone else None,
                "src": "nokhbegan_curated",
            }
        )
    for co, mgr, tech, phone in QUMS:
        hits.append(
            {
                "company": co,
                "person": f"{mgr} (مدیر — قزوین)",
                "mobile": None,
                "land": land(phone) if phone else None,
                "src": "qums_mgr",
            }
        )
        if tech:
            hits.append(
                {
                    "company": co,
                    "person": f"{tech} (مسئول فنی — قزوین)",
                    "mobile": None,
                    "land": land(phone) if phone else None,
                    "src": "qums_tech",
                }
            )
    for city, pe, phone, _ in CINA:
        hits.append(
            {
                "company": f"سینا پخش — مرکز {city}",
                "person": f"{pe} (مدیر مرکز سینا پخش {city})",
                "mobile": None,
                "land": land(phone),
                "src": "cina_curated",
            }
        )
    for co, pe, mob, ld in FARS:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیر انجمن فارس)",
                "mobile": to_mobile(mob),
                "land": land(ld) if ld else None,
                "src": "fars_assoc_snippet",
            }
        )
    for co, pe, phone in BIODEP:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیرعامل / نماینده استانی)",
                "mobile": None,
                "land": land(phone) if phone and len(re.sub(r'\D','',phone))>=10 else None,
                "src": "biodep",
            }
        )
    for co, pe, mob, ld in EXTRA:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیرعامل)",
                "mobile": to_mobile(mob),
                "land": land(ld),
                "src": "extra_ceo2",
            }
        )
    hits.extend(fetch_cina())
    hits.extend(fetch_fars_pdf())
    hits.extend(fetch_syndi())
    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
