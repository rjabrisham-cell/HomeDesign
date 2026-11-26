# -*- coding: utf-8 -*-
"""Import Behestan Pakhsh branch managers (company+person+landline)."""
from __future__ import annotations

import json
import re
from pathlib import Path

OUT = Path(r"D:\HomeDesign\research\distribution-contacts\batches\person_id_hits.jsonl")
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

# (city, person, role, landline)
ROWS = [
    ("اصفهان", "مجید موذن صفایی", "مدیر شعبه", "03136546212"),
    ("اصفهان", "حامد الفیده", "سرپرست فروش شعبه", "03136546212"),
    ("اصفهان", "سیدجعفر مهاجر موحد", "سرپرست فروش شعبه", "03136546212"),
    ("تهران غرب", "سعید هادیان", "سرپرست مرکز", "02144986547"),
    ("تهران غرب", "یداله محمدی", "سرپرست فروش مرکز", "02144986547"),
    ("تهران غرب", "حسین دشتی", "سرپرست فروش مرکز", "02144986547"),
    ("کرمان", "مهران معماریان", "مدیر شعبه", "03433350905"),
    ("کرمان", "امین کارگر", "سرپرست فروش شعبه", "03433350905"),
    ("فارس", "محسن یزدان‌خواه", "مدیر شعبه", "02184054982"),
    ("فارس", "محمد مختاری", "سرپرست فروش شعبه", "02184054982"),
    ("یزد", "محمد صالحی‌راد", "مدیر شعبه", ""),
    ("یزد", "محمدحسن غیاثی", "سرپرست فروش شعبه", ""),
    ("مازندران", "صابر اسمعیل‌زاده", "مدیر شعبه", "01133347530"),
    ("مازندران", "مهدی مهدی‌زاده", "سرپرست فروش شعبه", "01133347530"),
    ("همدان", "ابراهیم مرادی", "مدیر شعبه", "08132569237"),
    ("همدان", "امیر خدابنده", "سرپرست فروش شعبه", "08132569237"),
    ("مرکزی", "مجید باقری", "مدیر شعبه", "08634130325"),
    ("مرکزی", "مصطفی نراقی", "سرپرست فروش شعبه", "08634130325"),
]


def land(v: str) -> str | None:
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if not d or d.startswith("09"):
        return None
    if not d.startswith("0") and len(d) == 10:
        d = "0" + d
    return d if d.startswith("0") else None


def main() -> None:
    hits = []
    for city, pe, role, tel in ROWS:
        hits.append(
            {
                "company": f"بهستان پخش — شعبه {city}",
                "person": f"{pe} — {role} (behestanpakhsh.com)",
                "mobile": None,
                "land": land(tel),
                "src": "behestan_branches",
            }
        )
    existing = []
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if line.strip():
                existing.append(json.loads(line))
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            key = (h["company"], h["person"], "")
            if key in seen:
                continue
            seen.add(key)
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    print("appended", added)


if __name__ == "__main__":
    main()
