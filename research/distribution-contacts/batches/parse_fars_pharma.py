# -*- coding: utf-8 -*-
"""Parse Fars pharma member PDF into person_id_hits.jsonl (branch-aware rows)."""
from __future__ import annotations

import json
import re
from pathlib import Path

import fitz  # PyMuPDF

ROOT = Path(r"D:\HomeDesign")
PDF = ROOT / "research/distribution-contacts/raw/fars_pharma_members.pdf"
TXT = ROOT / "research/distribution-contacts/raw/fars_pharma_members.txt"
OUT = ROOT / "research/distribution-contacts/batches/person_id_hits.jsonl"
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

# Manual curated rows from OCR/search where parser may miss (company, person, mobile, land)
CURATED = [
    ("پخش هجرت — شعبه فارس", "حسن‌زاده — مدیر شعبه فارس", "09178871255", "07138247775"),
    ("سایه سمن — شعبه فارس", "بایزدان — مدیر شعبه فارس", "09178200919", "07137418682"),
    ("قاسم ایران — شعبه فارس", "کسرایی — مدیر شعبه فارس", "09171128581", "07137202449"),
    ("پخش البرز — شعبه فارس", "علیزاده — مدیر شعبه فارس", "09173000114", "07191010515"),
    ("پخش رازی — شعبه فارس", "زارع — مدیر شعبه فارس", "09171182965", "07132425316"),
    ("توزیع دارو پخش — شعبه فارس", "کشاورز — مدیر شعبه فارس", "09177169647", "07132268590"),
    ("بهستان پخش — شعبه فارس", "یزدان‌خواه — مدیر شعبه فارس", "09173031085", "02184054982"),
    ("پخش ممتاز — شعبه فارس", "نجفی — مدیر شعبه فارس", "09365826398", "07137425475"),
    ("آدوراطب — شعبه فارس", "زارعی — مدیر شعبه فارس", "09171125238", "07138221244"),
    ("پخش فردوس — شعبه فارس", "ترابیان — مدیر شعبه فارس", "09177025877", "07138424748"),
    ("محیا دارو — شعبه فارس", "حسینی راد — مدیر شعبه فارس", "09179327173", "07137742323"),
    ("پخش کارن — شعبه فارس", "سرگل‌زایی — مدیر شعبه فارس", "09128106295", "07137730016"),
    ("دارو سازان التیام — شعبه فارس", "صدر — مدیر شعبه فارس", "09173110024", "07137740343"),
    ("میلاد داروی فارس", "افشار — مدیرعامل", "09177878633", "07138228080"),
    ("داروگستر پیشرو پارسیان — شعبه فارس", "مهشید فروزش — مدیر شعبه فارس", "09173142377", "07132321002"),
    ("شرکت دارویی اکسیر — شعبه فارس", "سلطانی — مدیر شعبه فارس", "09179195250", "07138408387"),
    ("مشاء طب — شعبه فارس", "دهداری — مدیر شعبه فارس", "09171132280", ""),
    ("سها هلال — شعبه فارس", "بهادری — مدیر شعبه فارس", "09177154924", "07138424219"),
    ("باریج اسانس — شعبه فارس", "قاسم رکن‌پور — مدیر شعبه فارس", "09177052329", "07132426360"),
    ("الیت دارو — شعبه فارس", "شریفی‌نژاد — مدیر شعبه فارس", "09171062755", ""),
    ("سلامت پخش هستی — شعبه فارس", "سبزعلیان — مدیر شعبه فارس", "09171088556", "07137743258"),
    ("گران طب نوین — شعبه فارس", "احمدی — مدیر شعبه فارس", "09175388848", "07132267049"),
    # Isfahan board (co+pe)
    ("خاقانی نوین سپاهان", "جمشید سلطانی تهرانی — مدیرعامل / رئیس انجمن اصفهان", "", "03132327228"),
    ("پراور پخش سپاهان", "حسین همتی‌نژاد — مدیرعامل / نایب رئیس انجمن اصفهان", "", ""),
    ("دارو پخش — شعبه اصفهان", "مهدی علیخانی — مدیر شعبه اصفهان", "", ""),
    ("دارو بهداشت آسیا", "امیررضا مقاره‌عابد — رئیس هیئت مدیره", "", ""),
    ("مهر آیین آریا ایرانیان", "مجتبی جعفری — مدیرعامل", "", ""),
    ("پخش رازی — شعبه اصفهان", "مصطفی محمدی‌نژاد — مدیر شعبه اصفهان", "", ""),
    ("قاسم ایران — شعبه اصفهان", "محمود نظری — مدیر شعبه اصفهان", "", ""),
    ("کالا گستر پارتاک ایرانیان", "پدرام انصاری — مدیرعامل", "", ""),
    ("نگار پخش اشرافی", "جواد اشرافی بفرویی — رئیس هیئت مدیره", "", ""),
    ("بهداشت سلامت آوند فر — شعبه اصفهان", "غلامرضا جباری — مدیر شعبه اصفهان", "", ""),
    ("پخش آرمان مهر سلامت اوراسیا", "مهران خضری — مدیرعامل", "07145916748", "07137321695"),
]


def to_mobile(v: str) -> str | None:
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    if len(d) == 11 and d.startswith("09"):
        return d
    return None


def normalize_land(v: str) -> str | None:
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if not d or d.startswith("09") or (len(d) == 10 and d.startswith("9")):
        return None
    if d.startswith("98"):
        d = "0" + d[2:]
    if not d.startswith("0"):
        if len(d) == 8:
            d = "071" + d  # Shiraz default for this list
        elif len(d) == 10:
            d = "0" + d
    return d


def parse_pdf_blocks(text: str) -> list[dict]:
    """Heuristic block split on leading row numbers."""
    # Join spaced mobiles like "917 887 1255"
    text2 = re.sub(r"(?<!\d)(9\d{2})\s+(\d{3})\s+(\d{4})(?!\d)", r"\1\2\3", text)
    parts = re.split(r"(?m)^(?=\d{1,2}[\u0600-\u06FFA-Za-z])", text2)
    hits = []
    for part in parts:
        part = part.strip()
        if not part or not re.match(r"^\d{1,2}", part):
            continue
        mobiles = re.findall(r"0?9\d{9}", re.sub(r"\D", "", part) and part)
        # find mobiles more carefully in original part
        mob_cands = []
        for m in re.finditer(r"(?<!\d)(0?9\d{9})(?!\d)", re.sub(r"\s+", "", part)):
            mm = to_mobile(m.group(1))
            if mm:
                mob_cands.append(mm)
        # also spaced form already collapsed in text2
        for m in re.finditer(r"(?<!\d)(0?9\d{9})(?!\d)", text2[max(0, text2.find(part[:20])) :]):
            pass
        person = None
        m_person = re.search(
            r"(آقای|اقای|خانم|مهشید|دکتر)\s*([^\n\d]{2,40})", part
        )
        if m_person:
            person = (m_person.group(1) + " " + m_person.group(2)).strip()
            person = re.split(r"شیراز|کیلومتر|بلوار|شهرک|سراسری|استانی", person)[0].strip()
        # company: after row number until سراسری/استانی
        m_co = re.match(r"^\d{1,2}\s*([^\n]+?)(?:\n|سراسری|استانی)", part)
        company = None
        if m_co:
            company = m_co.group(1).strip()
        else:
            # fallback first line after number
            first = part.splitlines()[0]
            company = re.sub(r"^\d{1,2}", "", first).strip()
        if company and "؟؟؟" in (person or ""):
            continue
        if company and person and person not in ("؟؟؟",):
            hits.append(
                {
                    "company": f"{company} — شعبه فارس",
                    "person": f"{person} — مدیر/مدیرعامل شعبه فارس (adfars)",
                    "mobile": mob_cands[0] if mob_cands else None,
                    "land": None,
                    "src": "fars_pharma_pdf",
                }
            )
    return hits


def main() -> None:
    doc = fitz.open(PDF)
    text = "\n".join(page.get_text("text") for page in doc)
    TXT.write_text(text, encoding="utf-8")

    parsed = parse_pdf_blocks(text)
    print("parsed_blocks", len(parsed))

    # Prefer curated for quality; merge unique
    hits = []
    seen = set()
    for co, pe, mob, land in CURATED:
        key = (co, pe, to_mobile(mob) or "")
        if key in seen:
            continue
        seen.add(key)
        hits.append(
            {
                "company": co,
                "person": pe,
                "mobile": to_mobile(mob),
                "land": normalize_land(land) if land else None,
                "src": "fars_isfahan_curated",
            }
        )
    for h in parsed:
        key = (h["company"], h["person"], h.get("mobile") or "")
        if key in seen:
            continue
        seen.add(key)
        hits.append(h)

    # append (don't wipe prior hits if any non-fars)
    existing = []
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                o = json.loads(line)
            except json.JSONDecodeError:
                continue
            if o.get("src") in ("fars_pharma_pdf", "fars_isfahan_curated"):
                continue
            existing.append(o)

    with OUT.open("w", encoding="utf-8") as f:
        for h in existing + hits:
            f.write(json.dumps(h, ensure_ascii=False) + "\n")

    with_mob = sum(1 for h in hits if h.get("mobile"))
    print(json.dumps({"hits": len(hits), "with_mobile": with_mob}, ensure_ascii=False))


if __name__ == "__main__":
    main()
