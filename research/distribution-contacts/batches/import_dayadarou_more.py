# -*- coding: utf-8 -*-
"""Import Dayadarou branches + APIR/APEsf board people + leftover logiscm CEOs."""
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


def fetch(url: str, timeout: int = 30) -> str:
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")


def html_to_text(html: str) -> str:
    text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n+", "\n", text).translate(DIGITS)


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


DAYA = [
    ("تبریز", "09426002551", None),
    ("مازندران", None, "02191200210"),
    ("مشهد", None, "05136511084"),
    ("کرمانشاه", None, "08334735333"),
    ("زاهدان", None, "03491090491"),
    ("بجنورد", None, "05832412015"),
    ("کرمان", None, "03432726567"),
    ("ارومیه", None, "04432371122"),
    ("قم و اراک", "09426001668", None),
    ("همدان", None, "02191559800"),
    ("اردبیل", "09426024728", None),
    ("اهواز", "09426043623", None),
    ("اصفهان", "09426043621", None),
    ("رشت", "09426001043", None),
    ("زنجان", "09426001218", None),
    ("شیراز", None, "07191090787"),
    ("گرگان", None, "01791014029"),
    ("بندرعباس", None, "07691099541"),
    ("البرز", None, "02191093813"),
    ("خرم‌آباد", "09426002496", None),
    ("تهران شرق", None, "02191091269"),
    ("تهران غرب", None, "02191079646"),
    ("ایلام", None, "02191079343"),
    ("انبار مرکزی", None, "02191094704"),
    ("دفتر مرکزی", None, "02191200210"),
    ("پشتیبانی", None, "02192004658"),
]

# Alborz DC executives (co_pe)
ALBORZ = [
    ("جلال کرابی", "مدیرعامل پخش البرز", "02166741994"),
    ("علیرضا خدائی ساخلو", "رئیس هیئت مدیره پخش البرز", "02166741994"),
    ("میثم مفتاحی", "عضو هیئت مدیره پخش البرز", None),
    ("زهرا کاوه", "عضو هیئت مدیره پخش البرز", None),
    ("میثم سعیدی", "مدیر ارتباطات پخش البرز", None),
    ("سام زمانی", "مدیر فروش آنلاین پخش البرز", None),
]

# APEsf board from site
APESF = [
    ("جمشید سلطانی تهرانی", "رئیس هیئت مدیره / مدیرعامل خاقانی نوین سپاهان", "03132327228"),
    ("حسین همتی نژاد", "نایب رئیس / مدیرعامل پراور پخش سپاهان", "03132327228"),
    ("مهدی علیخانی", "خزانه‌دار / مدیر شعبه داروپخش اصفهان", "03132327228"),
    ("امیررضا مقاره عابد", "نایب رئیس هیئت مدیره انجمن اصفهان", "03132327228"),
    ("مجتبی جعفری", "عضو هیئت مدیره انجمن اصفهان", "03132327228"),
    ("مصطفی محمدی نژاد", "عضو هیئت مدیره انجمن اصفهان", "03132327228"),
    ("فرهاد دارابی", "عضو هیئت مدیره انجمن اصفهان", "03132327228"),
    ("مجید حیدری", "عضو علی‌البدل هیئت مدیره انجمن اصفهان", "03132327228"),
    ("مرتضی طهماسبی", "عضو علی‌البدل هیئت مدیره انجمن اصفهان", "03132327228"),
    ("جواد اشرافی", "بازرس انجمن اصفهان", "03132327228"),
    ("محمد کریمی", "بازرس علی‌البدل انجمن اصفهان", "03132327228"),
    ("حمیدرضا رادمهر", "دبیر انجمن اصفهان", "03132327228"),
]

# Extra CEOs from logiscm snippets / known lists
EXTRA_CEO = [
    ("داروگستر باریج اسانس", "ابوالفضل همتی", None, "02144279530"),
    ("داروگستر رازی", "دکتر پیروز شیخ رضایی", None, "02188537027"),
    ("داروگستر نخبگان", "بهزاد آسوده", None, "02187792000"),
    ("داروگستر یاسین", "محمد شیری", None, "02144968243"),
    ("دایا خوراک و بهداشت", "حمید سلامت", None, "02188782010"),
    ("دایا دارو", "دکتر شهیاد زمانی", None, "02144400940"),
    ("دایا دارو", "مهندس محمد صدرا ملکی", None, "02191200210"),
    ("دنیا پخش یزد", "علی موسوی", None, "02188748880"),
    ("دی دارو امید", "دکتر حسین نامور", None, "02188995975"),
    ("رهاگستر دنیا", "میثاق رزاقی داریان", None, "02188385065"),
    ("گلچین درخشان تهران", "سپیده عیوضی", None, "02188936082"),
    ("سلیم البرز", "سید هادی رئیس اسلام زاده", None, "01144833400"),
    ("پاک فارس", "دکتر علیرضا قدرتی", None, "07136324470"),
    ("زاگرس زرین پارس", "نادر فیروز", None, "07136324874"),
]


def scrape_apir_boards() -> list[dict]:
    hits = []
    # APIR branch pages typically /branches/branch-N
    for n in range(1, 35):
        url = f"https://apir.ir/branches/branch-{n}"
        try:
            html = fetch(url, timeout=15)
        except Exception:
            continue
        if "هیئت مدیره" not in html and "هییت مدیره" not in html and "نام و نام خانوادگی" not in html:
            continue
        text = html_to_text(html)
        title_m = re.search(r"انجمن\s+([^\n]{2,40})", text)
        province = title_m.group(1).strip() if title_m else f"شعبه {n}"
        # rows: name + role
        for m in re.finditer(
            r"(?:آقای|خانم)?\s*([آ-یA-Za-z\u0600-\u06FF]{2,}(?:\s+[آ-یA-Za-z\u0600-\u06FF]{2,}){1,4})\s*\n\s*(رئیس هیئت مدیره|نایب رئیس|خزانه دار|خزانه‌دار|عضو هیئت مدیره|عضو علی البدل|بازرس|دبیر[^\n]*)",
            text,
        ):
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
            role = m.group(2).strip()
            hits.append(
                {
                    "company": f"انجمن ملی صنعت پخش — {province}",
                    "person": f"{pe} ({role} انجمن {province})",
                    "mobile": None,
                    "land": None,
                    "src": "apir_board",
                    "url": url,
                }
            )
        # also capture landline from page
        lands = []
        for mm in re.finditer(r"(?<!\d)(0\d{2,3}\d{7,8})(?!\d)", text):
            d = mm.group(1)
            if not d.startswith("09") and d not in lands:
                lands.append(d)
        if lands and hits:
            # attach land to last few from this page without land
            for h in hits[-12:]:
                if h.get("url") == url and not h.get("land"):
                    h["land"] = lands[0]
    return hits


def main():
    hits: list[dict] = []

    # Dayadarou
    try:
        html = fetch("https://www.dayadarou.ir/branches")
        (RAW / "dayadarou_branches.html").write_text(html, encoding="utf-8")
    except Exception as e:
        print({"daya_fetch_error": str(e)})

    for city, mob, ld in DAYA:
        mb = to_mobile(mob)
        ln = land(ld)
        # company+mobile or company+person placeholder for land-only: use contact role
        pe = f"مخاطب مرکز دایادارو {city}"
        if mb or ln:
            hits.append(
                {
                    "company": f"دایا دارو — مرکز {city}",
                    "person": pe if not mb else None,  # prefer pure co_mob when mobile exists
                    "mobile": mb,
                    "land": ln,
                    "src": "dayadarou",
                }
            )
            # also add person+company when only land (co_pe)
            if not mb and ln:
                hits[-1]["person"] = pe

    for person, role, ld in ALBORZ:
        hits.append(
            {
                "company": "پخش البرز — دفتر مرکزی",
                "person": f"{person} ({role})",
                "mobile": None,
                "land": land(ld),
                "src": "alborz_exec",
            }
        )

    for person, role, ld in APESF:
        hits.append(
            {
                "company": "انجمن صنعت پخش اصفهان",
                "person": f"{person} ({role})",
                "mobile": None,
                "land": land(ld),
                "src": "apesf_board",
            }
        )

    for co, pe, mob, ld in EXTRA_CEO:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیرعامل)",
                "mobile": to_mobile(mob),
                "land": land(ld),
                "src": "logiscm_extra_ceo",
            }
        )

    hits.extend(scrape_apir_boards())
    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
