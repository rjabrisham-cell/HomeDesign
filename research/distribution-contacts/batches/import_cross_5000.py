# -*- coding: utf-8 -*-
"""Final push past 5000: full Shafaarad + HQ + leftover CEOs from logiscm dump."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "batches" / "person_id_hits.jsonl"

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


SHAFA = [
    ("تهران متابولیک", "احمدرضا جوادی", None),
    ("اصفهان", "محمد کریمی", "03135720750"),
    ("ارومیه", "هادی شیرین وند", "04491001331"),
    ("زاهدان", "مرتضی فقیهی", "05433592524"),
    ("تهران غرب", "تابان طالبی", "02144901534"),
    ("گرگان", "عبدالباسط یحیائی", "01734533873"),
    ("ساری", "حسین بهزادی", "01133685555"),
    ("رشت", "سعید رهبر", "01333835005"),
    ("یزد", "رضا بافنده", "03537275737"),
    ("شیراز", "عبدالرضا قاسم پور", "07132425902"),
    ("مشهد", "راستگو", "05136912412"),
    ("تبریز", "تقی پورولی", "04136372236"),
    ("اهواز", "سعید حسین نژادیان", "06132289955"),
    ("تهران شرق", "مجید انواری", "02192001315"),
    ("کرمان", "حسن فرح بخش", "03432711360"),
    ("همدان", "حبیب کرمی", "08134586771"),
    ("مرکزی", "عباسی", "02144906129"),
    ("قم", "نصراله کشاورزی", "02536641027"),
    ("قزوین", "علیرضا امینی", "02832849317"),
    ("کرمانشاه", "فرهاد اسفندیاری", "08334277574"),
]

SHAFA_HQ = [
    ("علیرضا خوز", "مدیرعامل شفا آراد"),
    ("سید محمد سبحانیان", "معاون مالی شفا آراد"),
    ("علی شاکری نژاد", "معاون بازرگانی و فروش شفا آراد"),
    ("هادی اسکندری", "معاون منابع انسانی و پشتیبانی شفا آراد"),
    ("آیت معصومی مجد", "مدیر ارشد مالی شفا آراد"),
    ("عباس کرم پور", "مدیر مالی شفا آراد"),
    ("علیرضا مهاجری", "مدیر پشتیبانی لجستیک شفا آراد"),
    ("عبدالرضا برزگر", "مدیر حقوقی شفا آراد"),
]

EXTRA = [
    ("محیا دارو", "حمید ترکاشوند", "02188468111"),
    ("پوراپخش", "محسن بیگلری", "02189318000"),
    ("مکتاف", "امیر گرمابدری", "02188317930"),
    ("رخ ارا پخش مهرآیین", "محسن بیگلری", "02161983801"),
    ("پخش محیا دارو", "حمید ترکاشوند", "02188468008"),
    ("پخش مکتاف", "امیر گرمابدری", "02188304486"),
]


def parse_logiscm_agent_file() -> list[dict]:
    hits = []
    base = Path(r"C:\Users\rjabr\.cursor\projects\D-HomeDesign\agent-tools")
    if not base.exists():
        return hits
    for path in base.glob("*.txt"):
        try:
            text = path.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        if "مدیر عامل" not in text:
            continue
        for m in re.finditer(
            r"\|\s*\d+\)?\s*\|\s*([^|]+)\|\s*([^|]*)\|\s*([^|]*)\|\s*([^|]*)\|\s*([^|]*)\|\s*([^|\n]+)\|",
            text,
        ):
            co = m.group(1).strip()
            phone = m.group(4).strip()
            pe = m.group(6).strip()
            if not co or not pe or pe in ("مدیر عامل", "---") or len(pe) < 3:
                continue
            d = re.sub(r"\D", "", phone.translate(DIGITS))
            if d.startswith("98"):
                d = "0" + d[2:]
            hits.append(
                {
                    "company": co,
                    "person": f"{pe} (مدیرعامل دایرکتوری لجستیک)",
                    "mobile": to_mobile(d),
                    "land": land(d),
                    "src": "logiscm_agent_parse",
                }
            )
    return hits


def main():
    hits: list[dict] = []
    for city, pe, phone in SHAFA:
        hits.append(
            {
                "company": f"شفا آراد — مرکز {city}",
                "person": f"{pe} (مدیر مرکز شفا آراد {city})",
                "mobile": None,
                "land": land(phone) if phone else None,
                "src": "shafa_full",
            }
        )
    for pe, role in SHAFA_HQ:
        hits.append(
            {
                "company": "شفا آراد — دفتر مرکزی",
                "person": f"{pe} ({role})",
                "mobile": None,
                "land": land("02143039000"),
                "src": "shafa_hq",
            }
        )
    for co, pe, phone in EXTRA:
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیرعامل)",
                "mobile": None,
                "land": land(phone),
                "src": "extra_final",
            }
        )
    dump = parse_logiscm_agent_file()
    print({"logiscm_parse": len(dump)})
    hits.extend(dump)
    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
