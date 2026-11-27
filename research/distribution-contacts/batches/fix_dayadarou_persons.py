# -*- coding: utf-8 -*-
import json
import re
from pathlib import Path

OUT = Path(__file__).resolve().parents[1] / "batches" / "person_id_hits.jsonl"
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

existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()]
seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
added = wm = 0
with OUT.open("a", encoding="utf-8") as f:
    for city, mob, ld in DAYA:
        h = {
            "company": f"دایا دارو — مرکز {city}",
            "person": f"مخاطب مرکز دایادارو {city}",
            "mobile": to_mobile(mob),
            "land": land(ld),
            "src": "dayadarou_fix",
        }
        key = (h["company"], h["person"], h.get("mobile") or "")
        if key in seen:
            continue
        seen.add(key)
        if h.get("mobile"):
            wm += 1
        f.write(json.dumps(h, ensure_ascii=False) + "\n")
        added += 1
print({"fixed_daya": added, "with_mobile": wm})
