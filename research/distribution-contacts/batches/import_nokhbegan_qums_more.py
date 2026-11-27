# -*- coding: utf-8 -*-
"""Import Nokhbegan branches, QUMS depot/provincial, Hejrat/YTS extras, logiscm dump."""
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


def first_land(phone: str):
    if not phone:
        return None
    for tok in re.split(r"[\s,/،\-]+", str(phone).translate(DIGITS)):
        # rebuild area-local if split badly; also accept concatenated
        pass
    # try full string digits sequences
    text = str(phone).translate(DIGITS)
    for m in re.finditer(r"0?\d{2,3}[\- ]?\d{6,8}", text):
        d = re.sub(r"\D", "", m.group(0))
        if not d.startswith("0"):
            # if looks like 28xxxxxxxx missing 0
            if len(d) in (10, 11):
                d = "0" + d
        ld = land(d)
        if ld:
            return ld
    d = re.sub(r"\D", "", text)
    if d.startswith("98"):
        d = "0" + d[2:]
    return land(d)


def fetch(url: str, timeout: int = 25) -> str:
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")


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


# Curated from nokhbegandc.com/Branch search snippet
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
    ("کرمان", "فرهاد سلمانی", "02187792000"),
    ("کرمانشاه", "حمید باتمانی", "08334281037"),
    ("مشهد", "احسان پیله ور", "05136511061"),
    ("اهواز", "مصطفی زاروی", "02187792000"),
    ("لرستان", "رضا جمالی", "06633435106"),
    ("ساری", "محمد زمان فرجی", None),
    ("قزوین", "روشن انصاری", "02832376617"),
]

# QUMS depot page
QUMS_DEPOT = [
    ("پخش فردوس — قزوین", "دباغ", "نیاقى", "02833292571"),
    ("پخش هجرت — قزوین", "خوبان", "فرامرز کارآگاه", "02833297591"),
    ("پخش التیام — قزوین", "کشاورز", "جمالی", "02833292252"),
    ("آدوراطب — قزوین", "شفاهی", "شاه تقی", "02832848155"),
    ("داروگستر نخبگان — قزوین", "انصاری", "میررحیمی", "02832376617"),
    ("قاسم ایران — قزوین", "مظفرپور", "فتحی", "02832580220"),
    ("دی دارو امید — قزوین", "نصیری", None, "02832224501"),
    ("داروگستر شایان اعتماد — قزوین", "مرادی", "سیاهکلی", "02833290831"),
]

# QUMS provincial managers + technical
QUMS_PROV = [
    ("داروپخش — قزوین", "میار صنمی", "بهناز نیاقی", None),
    ("فردوس — قزوین", "دباغ", "لیلا طیبی جزایری", None),
    ("هجرت — قزوین", "سلیمانی", "فرامرز کارآگاه", None),
    ("داروسازان التیام — قزوین", "زارعی", "صدیقه کلانتری", None),
    ("آدوراطب — قزوین", "شفاهی", "ثمین نظری", None),
    ("دارو گستر نخبگان — قزوین", "انصاری", "معصومه رفیع خواه", None),
    ("شفا آراد — قزوین", "امینی", "سایه ابراهیمی", None),
    ("قاسم ایران — قزوین", "نعمتی", "مرتضی حسینی", None),
    ("سینا پخش ژن — قزوین", "حسینی", "ناصر بابا خداوردی", None),
    ("بهستان پخش — قزوین", "کامیاب محمدی", "شایان وطن دوست", None),
    ("میلان پارس فارمد — قزوین", "کوشان", "یاسمن پور بابا", None),
    ("سایه سمن — قزوین", "رمضانی", "سینوش کوشان", None),
    ("سها سروش نوین آسمان — قزوین", "حسین لو", "زهرا اصغری ها", None),
    ("دهکدک — قزوین", "شعبانی", None, None),
]

EXTRA_CEO = [
    ("یاراطب ثامن", "سیاوش مروج تربتی", "مدیرعامل یاراطب ثامن", "02144083412"),
    ("یاراطب ثامن", "علیرضا شه پرست", "مدیرعامل سابق یاراطب ثامن", "02144083412"),
    ("پخش هجرت", "مسعود بهرامی نوید", "مدیرعامل پخش هجرت", "02166461091"),
    ("پخش هجرت", "علیرضا شه‌پرست", "مدیرعامل پخش هجرت (خبری)", "02166461091"),
    ("داروگستر نخبگان", "بهزاد آسوده", "مدیرعامل داروگستر نخبگان", "02187792000"),
]


def scrape_nokhbegan() -> list[dict]:
    hits = []
    try:
        html = fetch("https://www.nokhbegandc.com/Branch")
        (RAW / "nokhbegan_branch.html").write_text(html, encoding="utf-8")
        text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
        text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
        text = re.sub(r"<[^>]+>", "\n", text)
        text = re.sub(r"\n+", "\n", text).translate(DIGITS)
        (RAW / "nokhbegan_branch_text.txt").write_text(text, encoding="utf-8")
        # pattern: آقای NAME مدیر شعبه استان CITY
        for m in re.finditer(
            r"(?:آقای|خانم)?\s*([آ-یA-Za-z\u0600-\u06FF][آ-یA-Za-z\u0600-\u06FF\s‌]{3,40})\s*مدیر شعبه استان\s+([^\n]{2,40})",
            text,
        ):
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
            city = re.sub(r"\s+", " ", m.group(2)).strip()
            # nearby phone
            window = text[m.end() : m.end() + 250]
            ld = None
            mm = re.search(r"(0\d{2,3}\d{7,8})", window)
            if mm:
                ld = land(mm.group(1))
            hits.append(
                {
                    "company": f"داروگستر نخبگان — شعبه {city}",
                    "person": f"{pe} (مدیر شعبه نخبگان {city})",
                    "mobile": None,
                    "land": ld,
                    "src": "nokhbegan_scrape",
                }
            )
    except Exception as e:
        print({"nokhbegan_error": str(e)})
    return hits


def scrape_hejrat_centers() -> list[dict]:
    hits = []
    # try known API/json leftovers
    for name in ("hejrat_centers.json", "hejrat_centers2.json"):
        p = RAW / name
        if not p.exists():
            continue
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            continue
        items = data if isinstance(data, list) else data.get("centers") or data.get("data") or []
        for it in items:
            if not isinstance(it, dict):
                continue
            city = it.get("city") or it.get("name") or it.get("title") or "شعبه"
            pe = it.get("manager") or it.get("person") or it.get("ceo")
            phone = it.get("mobile") or it.get("phone") or it.get("tel")
            if pe:
                hits.append(
                    {
                        "company": f"پخش هجرت — {city}",
                        "person": f"{pe} (مدیر مرکز هجرت {city})",
                        "mobile": to_mobile(phone),
                        "land": first_land(phone) if phone else None,
                        "src": "hejrat_json",
                    }
                )
    # scrape site for center links
    try:
        html = fetch("https://www.hejratco.com/")
        (RAW / "hejrat_home.html").write_text(html, encoding="utf-8")
        links = set(re.findall(r'href="([^"]+)"', html))
        center_links = [l for l in links if any(k in l.lower() for k in ("center", "branch", "shoab", "مراکز", "مرکز"))]
        print({"hejrat_center_links": center_links[:20]})
    except Exception as e:
        print({"hejrat_home_error": str(e)})
    return hits


def parse_logiscm_dump() -> list[dict]:
    hits = []
    # agent-tools dump from websearch
    candidates = list(Path(r"C:\Users\rjabr\.cursor\projects\D-HomeDesign\agent-tools").glob("*.txt"))
    for path in candidates:
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        if "مدیر عامل" not in text and "مدیرعامل" not in text:
            continue
        # markdown table rows: | n) | company | addr | city | phone | email | ceo |
        for m in re.finditer(
            r"\|\s*\d+\)?\s*\|\s*([^|]+)\|\s*([^|]*)\|\s*([^|]*)\|\s*([^|]*)\|\s*([^|]*)\|\s*([^|]+)\|",
            text,
        ):
            co = m.group(1).strip()
            phone = m.group(4).strip()
            pe = m.group(6).strip()
            if not co or not pe or pe in ("مدیر عامل", "---"):
                continue
            if len(co) < 2 or len(pe) < 2:
                continue
            hits.append(
                {
                    "company": co,
                    "person": f"{pe} (مدیرعامل logiscm)",
                    "mobile": to_mobile(phone),
                    "land": first_land(phone),
                    "src": "logiscm_dump",
                }
            )
    return hits


def main():
    hits: list[dict] = []
    hits.extend(scrape_nokhbegan())
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
    for co, mgr, tech, phone in QUMS_DEPOT:
        hits.append(
            {
                "company": co,
                "person": f"{mgr} (مدیر شرکت — انبار قزوین)",
                "mobile": None,
                "land": land(phone) if phone else None,
                "src": "qums_depot",
            }
        )
        if tech:
            hits.append(
                {
                    "company": co,
                    "person": f"{tech} (مسئول فنی — انبار قزوین)",
                    "mobile": None,
                    "land": land(phone) if phone else None,
                    "src": "qums_depot_tech",
                }
            )
    for co, mgr, tech, phone in QUMS_PROV:
        hits.append(
            {
                "company": co,
                "person": f"{mgr} (مدیر/فروش — قزوین استانی)",
                "mobile": None,
                "land": land(phone) if phone else None,
                "src": "qums_prov",
            }
        )
        if tech:
            hits.append(
                {
                    "company": co,
                    "person": f"{tech} (مسئول فنی — قزوین استانی)",
                    "mobile": None,
                    "land": None,
                    "src": "qums_prov_tech",
                }
            )
    for co, pe, role, phone in EXTRA_CEO:
        hits.append(
            {
                "company": co,
                "person": f"{pe} ({role})",
                "mobile": None,
                "land": land(phone),
                "src": "extra_ceo",
            }
        )
    hits.extend(scrape_hejrat_centers())
    dump = parse_logiscm_dump()
    print({"logiscm_dump_rows": len(dump)})
    hits.extend(dump)
    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
