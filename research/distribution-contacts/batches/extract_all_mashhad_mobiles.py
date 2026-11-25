# -*- coding: utf-8 -*-
"""Mobile-first extraction from Mashhad PDF text → person_id_hits.jsonl."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(r"D:\HomeDesign")
TXT = ROOT / "research/distribution-contacts/raw/mashhad_pharma.txt"
OUT = ROOT / "research/distribution-contacts/batches/person_id_hits.jsonl"
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

NOISE = re.compile(
    r"نام شرکت|نام مدير|نام مدیر|تلفن ثابت|تلفن همراه|نام مسئول|آدرس شرکت|مجوز|داروی مخدر|جوز دريافت|پوشش استان"
)


def to_mobile(v: str) -> str | None:
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    if len(d) == 11 and d.startswith("09"):
        return d
    return None


def main() -> None:
    raw = TXT.read_text(encoding="utf-8").translate(DIGITS)
    # flatten to single string with spaces
    flat = re.sub(r"\s+", " ", raw)

    hits = []
    for m in re.finditer(r"(0?9\d{9})", flat):
        mob = to_mobile(m.group(1))
        if not mob:
            continue
        start = max(0, m.start() - 80)
        window = flat[start : m.start()]
        # strip trailing landline digits stuck before mobile
        window = re.sub(r"[\d\-\s/]{5,}$", "", window).strip()
        # person = last Persian name chunk
        pe_m = re.search(
            r"((?:آقای|اقای|خانم|دکتر|مهندس)?\s*[آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهیءؤئيك‌\.]{2,}(?:\s+[آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهیءؤئيك‌\.]{2,}){0,3})\s*$",
            window,
        )
        if not pe_m:
            continue
        person = re.sub(r"\s+", " ", pe_m.group(1)).strip(" -–—.")
        if len(person) < 3 or NOISE.search(person):
            continue
        # company: look further back
        before = window[: pe_m.start()].strip()
        before = NOISE.sub(" ", before)
        before = re.sub(r"[√X×x]+", " ", before)
        # take last company-like token sequence
        co_m = re.search(
            r"([آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهیءؤئيك‌\s\(\)\.]{3,60})$",
            before,
        )
        company = None
        if co_m:
            company = re.sub(r"\s+", " ", co_m.group(1)).strip(" -–—.")
            # trim address prefixes
            company = re.split(
                r"(?:آزادی|بلوار|خیابان|کیلومتر|شهرک|توس\d|منطقه|طرق|بزرگراه|جاده)",
                company,
            )[-1].strip(" -–—.")
        if not company or len(company) < 3:
            company = "پخش دارویی مشهد (نام شرکت نامشخص)"
        if NOISE.search(company) or company.startswith("آزادی"):
            continue
        hits.append(
            {
                "company": f"{company} — شعبه مشهد",
                "person": f"{person} — مخاطب پخش (فهرست سلامت مشهد)",
                "mobile": mob,
                "land": None,
                "src": "mashhad_mobile_first",
            }
        )

    # dedupe by mobile
    by_mob = {}
    for h in hits:
        by_mob[h["mobile"]] = h
    hits = list(by_mob.values())

    existing = []
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if line.strip():
                existing.append(json.loads(line))
    seen_mobs = {h.get("mobile") for h in existing if h.get("mobile")}
    added = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            if h["mobile"] in seen_mobs:
                continue
            seen_mobs.add(h["mobile"])
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    print(json.dumps({"extracted_unique_mobiles": len(hits), "appended": added}, ensure_ascii=False))


if __name__ == "__main__":
    main()
