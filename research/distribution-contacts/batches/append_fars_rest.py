# -*- coding: utf-8 -*-
"""Append remaining Fars pharma rows (25-49) + Isfahan extras into person_id_hits.jsonl."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(r"D:\HomeDesign")
OUT = ROOT / "research/distribution-contacts/batches/person_id_hits.jsonl"
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

ROWS = [
    ("داروگستر طوبی جهرم", "کشافی — مدیر شعبه/مدیرعامل", "09173916084", "07154335192"),
    ("سینا پخش ژن — شعبه فارس", "علوی — مدیر شعبه فارس", "09178060105", "07137742957"),
    ("دارو جویان سرآمد فارس", "باصری — مدیرعامل", "09177055474", "07132341392"),
    ("یارا طب ثامن — شعبه فارس", "طاهری — مدیر شعبه فارس", "09176064479", "07137743705"),
    ("اهورا پخش امین — شعبه فارس", "عرب — مدیر شعبه فارس", "09171103030", "07132263210"),
    ("شفا آراد — شعبه فارس", "صداقت‌پیشه — مدیر شعبه فارس", "09177112245", "07132425902"),
    ("تجهیز گستر پاسارگاد جهرم", "فرح‌آبادی — مدیرعامل", "09177927976", "07154447005"),
    ("دارو گستر نخبگان — شعبه فارس", "اسماعیلی — مدیر شعبه فارس", "09177034802", "07132426522"),
    ("مهبان دارو — شعبه فارس", "بانشی — مدیر شعبه فارس", "09177053415", "07137441163"),
    ("سبز پویان طبیعت دراک", "کاظمی — مدیرعامل", "09173395575", "07138253320"),
    ("شیمی درمان اوسینا — شعبه فارس", "سلطانی — مدیر شعبه فارس", "09178559890", ""),
    ("آراد سلامت بوعلی — شعبه فارس", "تعجب — مدیر شعبه فارس", "09170101021", "07137743385"),
    ("درمان یاب پخش مهر آریا — شعبه فارس", "زارع — مدیر شعبه فارس", "09178845207", "07137744480"),
    ("بهداشت سلامت آوندفر — شعبه فارس", "سلجوقیان‌راد — مدیر شعبه فارس", "09171179129", "07136704370"),
    ("مشکات فارمد توزیع — شعبه فارس", "صفری — مدیر شعبه فارس", "09179260050", "07132425426"),
    ("میلان پارس فارمد — شعبه فارس", "داورپناه — مدیر شعبه فارس", "09171092892", "07137744256"),
    ("دارویی نیک فردای سالم (پخش نفس) — شعبه فارس", "کشتکاران — مدیر شعبه فارس", "09171003605", "07137425475"),
    ("آرمان سلامت فارابی — شعبه فارس", "کریمی — مدیر شعبه فارس", "09176361727", ""),
    ("داروگستر دراک پارس", "قبادی آزاد — مدیرعامل", "09173054419", ""),
    ("آرتا پخش آداک تدبیر — شعبه فارس", "یاسمیان — مدیر شعبه فارس", "09387263711", ""),
    ("دی دارو امید — شعبه فارس", "محمدصادق غلامپور — مدیر شعبه فارس", "09177182764", ""),
    ("شایا پخش ویرا — شعبه فارس", "بیژن صدر — مدیر شعبه فارس", "09176331002", ""),
    ("نوش دارو — شعبه فارس", "حسینی — مدیر شعبه فارس", "09177871948", ""),
    # Fars association board people as co+pe when company known
    ("انجمن صنعت پخش فارس", "عبدالرضا دیداری — رئیس هیئت مدیره", "", "07138350807"),
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
            d = "071" + d
        elif len(d) == 10:
            d = "0" + d
    return d


def main() -> None:
    existing = []
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if line.strip():
                existing.append(json.loads(line))
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = 0
    with OUT.open("a", encoding="utf-8") as f:
        for co, pe, mob, land in ROWS:
            h = {
                "company": co,
                "person": pe,
                "mobile": to_mobile(mob),
                "land": normalize_land(land) if land else None,
                "src": "fars_rest",
            }
            key = (h["company"], h["person"], h["mobile"] or "")
            if key in seen:
                continue
            seen.add(key)
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    print("appended", added)


if __name__ == "__main__":
    main()
