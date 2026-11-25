# -*- coding: utf-8 -*-
"""Parse Mashhad pharma distributor PDF into person_id_hits.jsonl."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(r"D:\HomeDesign")
TXT = ROOT / "research/distribution-contacts/raw/mashhad_pharma.txt"
OUT = ROOT / "research/distribution-contacts/batches/person_id_hits.jsonl"
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

SKIP = {
    "نام شرکت",
    "نام مدير شرکت",
    "نام مدیر شرکت",
    "تلفن ثابت",
    "تلفن همراه",
    "نام مسئول فني",
    "نام مسئول فنی",
    "آدرس شرکت",
    "مجوز زنجيره سرد",
    "مجوز زنجیره سرد",
    "داروی مخدر",
    "جوز دريافت داروهای شيمي درما",
    "مجوز دريافت سرم",
    "مجوز دریافت سرم",
    "مجوز پوشش استان های همجوار",
}


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
            d = "051" + d
        elif len(d) == 7:
            d = "051" + d
        elif len(d) == 10:
            d = "0" + d
    return d


def clean_name(s: str) -> str:
    s = S(s)
    s = re.sub(r"^\(?اعلم داروی مجد\)?", "", s)
    s = re.sub(r"[√X×x]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip(" -–—\t")
    return s


def S(v) -> str:
    return "" if v is None else str(v).strip()


def extract_person_mobile_pairs(blob: str) -> list[tuple[str, str]]:
    """Pull (person, mobile) from concatenated OCR like 'خانم حجازی9151197270دکتر ...'."""
    pairs = []
    # normalize: insert split before doctor/آقای after digits
    blob = blob.translate(DIGITS)
    # find all mobiles with preceding name fragment
    for m in re.finditer(
        r"([آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهیءؤئيك‌\s\.]{2,40}?)(0?9\d{9})",
        blob,
    ):
        name = clean_name(m.group(1))
        # strip trailing landline digits stuck to name
        name = re.sub(r"\d{5,}$", "", name).strip()
        name = re.sub(r"^(آقای|اقای|خانم|دکتر|مهندس)\s*", r"\1 ", name).strip()
        mob = to_mobile(m.group(2))
        if not mob or len(name) < 3:
            continue
        if name in SKIP or name.startswith("آزادی") or name.startswith("بلوار"):
            continue
        # reject if name is mostly digits/noise
        if sum(ch.isdigit() for ch in name) > 2:
            continue
        pairs.append((name, mob))
    return pairs


def main() -> None:
    text = TXT.read_text(encoding="utf-8")
    # collapse whitespace for pattern matching but keep newlines for company lines
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    hits: list[dict] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if line in SKIP or re.fullmatch(r"[√X×x]+", line):
            i += 1
            continue
        # company candidates: Persian-heavy, no long digit run at start
        if re.search(r"[آ-ی]", line) and not re.match(r"^\d", line) and line not in SKIP:
            # gather following blob until next likely company or header
            blob_parts = []
            j = i + 1
            while j < len(lines) and j < i + 6:
                nxt = lines[j]
                if nxt in SKIP:
                    break
                # stop if next looks like a new company (Persian, short, no mobile glued)
                if (
                    re.search(r"[آ-ی]", nxt)
                    and not re.search(r"0?9\d{9}", nxt.translate(DIGITS))
                    and not re.search(r"\d{7,}", nxt.translate(DIGITS))
                    and len(nxt) < 50
                    and j > i + 1
                ):
                    break
                blob_parts.append(nxt)
                j += 1
            blob = " ".join(blob_parts)
            pairs = extract_person_mobile_pairs(blob)
            lands = []
            for m in re.finditer(r"(?<!\d)(\d{7,11})(?!\d)", blob.translate(DIGITS)):
                land = normalize_land(m.group(1))
                if land and land not in lands:
                    lands.append(land)
            if pairs:
                # first pair = company manager; rest may be technical managers — keep all as valuable contacts
                for idx, (pe, mob) in enumerate(pairs):
                    role = "مدیر شرکت" if idx == 0 else "مسئول فنی/مخاطب"
                    hits.append(
                        {
                            "company": f"{clean_name(line)} — شعبه مشهد",
                            "person": f"{pe} — {role} (فهرست پخش سلامت مشهد)",
                            "mobile": mob,
                            "land": lands[0] if lands else None,
                            "src": "mashhad_pharma_pdf",
                        }
                    )
                i = j
                continue
        i += 1

    # also hard-code high-confidence rows from OCR samples / search snippets
    curated = [
        ("ارشیا دارو نوین ایرانیان — شعبه مشهد", "حجازی — مدیر شرکت", "09151197270", "05136500000"),
        ("دلوفان دارو (اعلم داروی مجد) — شعبه مشهد", "گلی — مدیر شرکت", "09399161666", ""),
        ("افق داروی پارس — شعبه مشهد", "حبیب‌الله جلینی — مدیر شرکت", "09155574149", "05136576174"),
        ("اکسیر — شعبه مشهد", "حمیدرضا بابازاده — مدیر شرکت", "09153101746", "05136220166"),
        ("البرز — شعبه مشهد", "عزیزاله امیری — مدیر شرکت", "09153018470", "05136571940"),
        ("الیت دارو — شعبه مشهد", "محمد صیامی — مدیر شرکت", "09124309204", "05136147370"),
        ("اهورا پخش امین — شعبه مشهد", "کاوه — مدیر شرکت", "09368097933", "05136512764"),
        ("آدورا طب — شعبه مشهد", "عادل — مدیر شرکت", "09158010818", "05135412068"),
        ("آرتا پخش آداک تدبیر — شعبه مشهد", "روحبخش — مدیر شرکت", "09151097418", ""),
        ("آریا فارمد صادقیان — شعبه مشهد", "صادقیان — مدیر شرکت", "09151016609", ""),
        ("آوا دارو تدبیر شرق — شعبه مشهد", "مراتی — مدیر شرکت", "09157894164", ""),
        ("برگ سبز آترین — شعبه مشهد", "قویدل — مدیر شرکت", "09151004363", ""),
        ("به پخش داروی بهشهر — شعبه مشهد", "افتخاری — مدیر شرکت", "09155151189", "05133920230"),
        ("بهداشت سلامت آوندفر — شعبه مشهد", "ایمن — مدیر شرکت", "09151115045", "05138812200"),
        ("بهرسان دارو — شعبه مشهد", "مجید خالقی — مدیر شرکت", "09128508656", "05136515140"),
        ("بهستان پخش — شعبه مشهد", "وحیدرضا بهاری — مدیر شرکت", "09151159140", "05136512509"),
        ("بهشاد آراء شرق مدهوشی — شعبه مشهد", "مدهوشی — مدیر شرکت", "09153164813", ""),
        ("پورا پخش — شعبه مشهد", "هادی وفادار — مدیر شرکت", "09151117211", "05136513420"),
        ("پونل برسام — شعبه مشهد", "سینا نوید عالیی — مدیر شرکت", "09121329810", "05138830798"),
        ("پویش فارمد گستر — شعبه مشهد", "ریحانه اسدالهی‌پور — مدیر شرکت", "09151236232", "05136512552"),
        ("گلرنگ پخش", "پیمان کاتبی — مدیرعامل (rasmio/linkedin)", "", "02184237000"),
        ("پخش پاکنام", "ناصر راسخی — مدیرعامل (rasmio)", "", ""),
        ("حافظ پخش", "محمدتقی براتی — مدیرعامل (rasmio)", "", ""),
        ("گروه پخش انسی", "مهیار رفیعی چیرانی — مدیرعامل (rasmio/logiscm)", "", "02144287623"),
        ("پخش پارس خزر", "طه میرعمادی — مدیرعامل (rasmio)", "", ""),
        ("قاسم ایران", "غلامعباس درفشی — مدیرعامل (سایت قاسم ایران)", "", "02144543200"),
        ("اکسیر", "محمدهادی کمالی — مدیرعامل (syndipharma)", "", "02177687801"),
        ("پخش البرز", "عقیل آرین‌نژاد — مدیرعامل (syndipharma)", "", "02166702435"),
        ("توزیع دارو پخش", "بیژن کلانتری‌اشان — مدیرعامل (syndipharma)", "", "02188888073"),
        ("پخش رازی", "محمدرضا اسدآبادی — مدیرعامل (syndipharma/logiscm)", "", "02166715712"),
        ("پخش فردوس", "مجید زجاجی — مدیرعامل (syndipharma)", "", "02166035170"),
        ("پخش هجرت", "مسعود بهرامی نوید — مدیرعامل (syndipharma)", "", "02166461933"),
        ("مهبان دارو", "علی باقری — مدیرعامل (syndipharma)", "", "02144626506"),
    ]
    for co, pe, mob, land in curated:
        hits.append(
            {
                "company": co,
                "person": pe,
                "mobile": to_mobile(mob),
                "land": normalize_land(land) if land else None,
                "src": "mashhad_curated",
            }
        )

    # merge into jsonl
    existing = []
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if line.strip():
                existing.append(json.loads(line))
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            key = (h.get("company"), h.get("person"), h.get("mobile") or "")
            if key in seen or not h.get("person"):
                continue
            seen.add(key)
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    with_mob = sum(1 for h in hits if h.get("mobile"))
    print(json.dumps({"parsed_or_curated": len(hits), "appended": added, "with_mobile_in_batch": with_mob}, ensure_ascii=False))


if __name__ == "__main__":
    main()
