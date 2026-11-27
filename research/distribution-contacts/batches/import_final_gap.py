# -*- coding: utf-8 -*-
"""Close remaining KPI gap: Shafaarad, Exir leftovers, Cina more, Imenpakhsh."""
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
    ("گیلان", "سعید رهبر", "01333835005"),
    ("یزد", "رضا بافنده", "03537275737"),
    ("کرمان", "حسن فرح بخش", "03432711360"),
    ("همدان", "حبیب کرمی", "08134586771"),
    ("مرکزی", "عباسی", "02144906129"),
    ("کرمانشاه", "فرهاد اسفندیاری", "08334277574"),
]

EXIR = [
    ("قم", "محمد وزینی", None, "02536642220"),
    ("کرمان", "حسنعلی بلالی", "09138909010", "03432752944"),
    ("کرمانشاه", "اسماعیل علم‌خواه", "09128953274", "08338353011"),
    ("گرگان", "کریم کوثری", "09112718066", "01732170934"),
    ("رشت", "ایمان قربان پور", "09118274048", "01331883183"),
    ("یزد", "محمود تقیانی", None, None),
    ("مازندران", "امید عباس زاده", None, None),
]

CINA_MORE = [
    ("اهواز", "محمود مقامی حیدر", "06133908800"),
    ("رشت", "حمید قاهری شرقی", "01331884691"),
    ("اصفهان", "جواد نقیان", "03135721948"),
    ("یزد", "مرتضایی پور فرهاد", "03537272250"),
    ("کرمانشاه", "امیرحسین مهرابی", "08334735216"),
    ("همدان", "حمیدرضا مقصود", "08134384051"),
    ("اراک", "ایمان رحیمی", "08634173351"),
]

ELTIAM_MORE = [
    ("همدان", "صابر صید محمدی", "08134383795"),
    ("گیلان", "علی خسروی", "01333696261"),
    ("کرمان", "محبوبه شکیبا نیا", "03432719595"),
    ("یزد", "مهدی ادهمی", "03535219544"),
    ("کرمانشاه", "علیرضا علائی", "08334269766"),
    ("سنندج", "محمد سامان مرب", "08733842089"),
]

IMEN = [
    ("آگرین پخش سلامت — کرمانشاه", None, "08334324473"),
    ("پیشگامان بهداشت گستر غرب — کرمانشاه", None, "08334268980"),
    ("ایمن پارس نورا — گیلان", None, "01333115236"),
    ("ایمن سازان طیور شمال — گیلان", None, "01333594011"),
    ("دام گستر لرستان", None, "06642610347"),
    ("پارس مهر پیوند — مازندران", None, "01144235050"),
    ("آمار دارو داران — مازندران", None, "01144221051"),
    ("طب گستر کاوه — مرکزی", None, "08634130918"),
    ("جهان طب گستر ایرانیان", None, "02144036322"),
    ("هستی پخش هرمزگان", None, "07152252938"),
    ("دامین گستر هگمتان — همدان", None, "08133263469"),
    ("ایمن پخش نهاوند", None, "08133264133"),
    ("پرنیان طب ایساتیس — یزد", None, "03538272689"),
    ("صحت دام یزد", None, "03537223380"),
    ("دام یاوران یزد", None, "03538254479"),
]


def scrape_shafa() -> list[dict]:
    hits = []
    try:
        req = urllib.request.Request("https://www.shafaarad.ir/branch", headers=UA)
        html = urllib.request.urlopen(req, timeout=20).read().decode("utf-8", "ignore")
        (RAW / "shafaarad_branch.html").write_text(html, encoding="utf-8")
        text = re.sub(r"<[^>]+>", "\n", html)
        text = re.sub(r"\n+", "\n", text).translate(DIGITS)
        for m in re.finditer(
            r"نام مدیر\s*:\s*(?:آقای|خانم)?\s*([^\n]+)\s*\n\s*شماره تماس\s*:\s*([^\n]+)",
            text,
        ):
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
            phone = re.sub(r"\s+", " ", m.group(2)).strip()
            # find preceding heading-ish city
            prev = text[max(0, m.start() - 120) : m.start()]
            cm = re.findall(r"\n\s*([آ-یA-Za-z\u0600-\u06FF][^\n]{1,40})\s*\n", prev)
            city = cm[-1].strip() if cm else "شعبه"
            hits.append(
                {
                    "company": f"شفا آراد — مرکز {city}",
                    "person": f"{pe} (مدیر مرکز شفا آراد {city})",
                    "mobile": to_mobile(phone),
                    "land": land(phone if phone.startswith("0") else "0" + re.sub(r"\D", "", phone)),
                    "src": "shafa_scrape",
                }
            )
        print({"shafa_scrape": len(hits)})
    except Exception as e:
        print({"shafa_error": str(e)})
    return hits


def main():
    hits: list[dict] = []
    hits.extend(scrape_shafa())
    for city, pe, phone in SHAFA:
        hits.append(
            {
                "company": f"شفا آراد — مرکز {city}",
                "person": f"{pe} (مدیر مرکز شفا آراد {city})",
                "mobile": None,
                "land": land(phone),
                "src": "shafa_curated",
            }
        )
    for city, pe, mob, ld in EXIR:
        hits.append(
            {
                "company": f"پخش اکسیر — مرکز {city}",
                "person": f"{pe} (مدیر مرکز اکسیر {city})",
                "mobile": to_mobile(mob),
                "land": land(ld) if ld else None,
                "src": "exir_gap",
            }
        )
    for city, pe, phone in CINA_MORE:
        hits.append(
            {
                "company": f"سینا پخش — مرکز {city}",
                "person": f"{pe} (مدیر مرکز سینا پخش {city})",
                "mobile": None,
                "land": land(phone),
                "src": "cina_more",
            }
        )
    for city, pe, phone in ELTIAM_MORE:
        hits.append(
            {
                "company": f"داروسازان التیام — مرکز {city}",
                "person": f"{pe} (مدیر مرکز التیام {city})",
                "mobile": None,
                "land": land(phone),
                "src": "eltiam_more",
            }
        )
    for co, pe, phone in IMEN:
        hits.append(
            {
                "company": co,
                "person": pe or f"مخاطب نمایندگی {co.split('—')[0].strip()}",
                "mobile": None,
                "land": land(phone),
                "src": "imenpakhsh",
            }
        )
    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
