# -*- coding: utf-8 -*-
"""Import Fars association members (with mobiles) + remaining Exir centers."""
from __future__ import annotations

import json
import re
from pathlib import Path

OUT = Path(r"D:\HomeDesign\research\distribution-contacts\batches\person_id_hits.jsonl")
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
    return d if d.startswith("0") and 10 <= len(d) <= 11 else None


# Fars association PDF (adfars) — company, manager, mobile, land
FARS = [
    ("پخش هجرت — شیراز", "حسن‌زاده — مدیر (انجمن فارس)", "09178871255", "07138247775"),
    ("سایه سمن — شیراز", "بایزدان — مدیر (انجمن فارس)", "09178200919", "07137418682"),
    ("قاسم ایران — شیراز", "کسرایی — مدیر (انجمن فارس)", "09171128581", "07137202449"),
    ("قاسم ایران — شیراز", "مخاطب دوم (انجمن فارس)", "09173101287", "07137202449"),
    ("پخش البرز — شیراز", "علیزاده — مدیر (انجمن فارس)", "09173000114", "07191010515"),
    ("پخش رازی — شیراز", "زارع — مدیر (انجمن فارس)", "09171182965", "07132425316"),
    ("توزیع داروپخش — شیراز", "کشاورز — مدیر (انجمن فارس)", "09177169647", "07132268590"),
    ("بهستان پخش — شیراز", "یزدانخواه — مدیر (انجمن فارس)", "09173031085", "02184054982"),
    ("پخش ممتاز — شیراز", "نجفی — مدیر (انجمن فارس)", "09365826398", "07137425475"),
    ("آدورا طب — شیراز", "زارعی — مدیر (انجمن فارس)", "09171125238", "07138221244"),
    ("پخش فردوس — شیراز", "ترابیان — مدیر (انجمن فارس)", "09177025877", "07138424748"),
    ("محیا دارو — شیراز", "حسینی — مدیر (انجمن فارس)", "09179327173", "07137742323"),
    ("پخش کارن — شیراز", "سرگل‌زائی — مدیر (انجمن فارس)", "09128106295", "07137730016"),
    ("داروسازان التیام — شیراز", "صدر — مدیر (انجمن فارس)", "09173110024", "07137740341"),
    ("میلاد داروی فارس — شیراز", "افشار — مدیر (انجمن فارس)", "09177878633", "07138228080"),
    ("میلاد داروی فارس — شیراز", "مخاطب دوم (انجمن فارس)", "09178961311", "07138228080"),
    ("داروگستر پیشرو پارسیان — شیراز", "مهشید — مدیر (انجمن فارس)", "09173142377", "07132321002"),
    ("داروگستر پیشرو پارسیان — شیراز", "موسوی — مخاطب (انجمن فارس)", "09178011803", "07132321002"),
    ("پخش اکسیر — شیراز", "سلطانی — مدیر (انجمن فارس)", "09179195250", "07138408387"),
    ("مشاء طب — شیراز", "دهداری — مدیر (انجمن فارس)", "09171132280", None),
    ("سها هلال احمر — شیراز", "مدیر شعبه (انجمن فارس)", "09177154924", "07138424219"),
    ("سها هلال احمر — شیراز", "بهادری — مخاطب (انجمن فارس)", "09395624255", "07138424219"),
    ("الیت دارو — شیراز", "شریفی‌نژاد — مدیر (انجمن فارس)", "09171062755", None),
    ("سلامت پخش هستی — شیراز", "سبزعلیان — مدیر (انجمن فارس)", "09171088556", "07137743258"),
    ("تجهیزگران طب نوین — شیراز", "احمدی — مدیر (انجمن فارس)", "09175388488", "07132267049"),
]

EXIR_MORE = [
    ("بابل/مازندران", "امید عباس‌زاده", None, "01132078159"),
    ("بندرعباس", "فاطمه روشن", None, "07632561109"),
    ("یزد", "محمود تقیانی", None, None),
    ("بروجرد", "مدیر مرکز بروجرد", None, "06642469345"),
]

EXIR_BOARD = [
    ("پخش اکسیر", "حمید قربانی — رئیس هیئت مدیره (exirpd)"),
    ("پخش اکسیر", "محمدرضا حسن‌زاده قصبه — عضو هیئت مدیره (exirpd)"),
]


def main() -> None:
    hits: list[dict] = []
    for co, pe, mob, tel in FARS:
        hits.append(
            {
                "company": co,
                "person": pe,
                "mobile": to_mobile(mob),
                "land": land(tel) if tel else None,
                "src": "adfars_assoc",
            }
        )
    for city, pe, mob, tel in EXIR_MORE:
        hits.append(
            {
                "company": f"پخش اکسیر — مرکز {city}",
                "person": f"{pe} — مدیر مرکز (exirpd)",
                "mobile": to_mobile(mob) if mob else None,
                "land": land(tel) if tel else None,
                "src": "exir_more",
            }
        )
    for co, pe in EXIR_BOARD:
        hits.append({"company": co, "person": pe, "mobile": None, "land": None, "src": "exir_board"})

    existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()]
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = 0
    with_mob = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            key = (h.get("company"), h.get("person"), h.get("mobile") or "")
            if key in seen:
                continue
            seen.add(key)
            if h.get("mobile"):
                with_mob += 1
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    print(json.dumps({"batch": len(hits), "appended": added, "with_mobile": with_mob}, ensure_ascii=False))


if __name__ == "__main__":
    main()
