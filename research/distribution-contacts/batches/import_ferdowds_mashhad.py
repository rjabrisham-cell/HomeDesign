# -*- coding: utf-8 -*-
"""Import Ferdows nationwide centers + Mashhad FDO PDF hits with mobiles."""
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
    if not d.startswith("0"):
        if len(d) == 10:
            d = "0" + d
        elif len(d) == 8:  # Tehran local
            d = "021" + d
        elif len(d) == 7:
            return None
    return d if d.startswith("0") and 10 <= len(d) <= 11 else None


# Ferdows warehouses: (city, manager, phone)
FERDOWS = [
    ("تهران", "محسن کریمی", "02144542921"),
    ("تبریز", "رحیم زارع", "04136377500"),
    ("رشت", "کامبیز سرمست", "01333660001"),
    ("ساری", "علیرضا یوسف‌داد", "01133281220"),
    ("مشهد", "وحید سجادی", "05133920163"),
    ("کرمان", "محمد معاذالهی", "03432163966"),
    ("شیراز", "آرش ترابیان", "07138225633"),
    ("اصفهان", "متین صیامی", "03136548444"),
    ("اهواز", "محمدرضا حیدرپور", "06134486023"),
    ("کرمانشاه", "مریم قلی‌نژاد", "08338266489"),
    ("بندرعباس", "یاسر ریگی", "07632561318"),
    ("ارومیه", "کیوان حلوائی‌نژاد", "04432376640"),
    ("اراک", "علیرضا اسدی", "08633134043"),
    ("زاهدان", "محمدرضا حسین‌پور", "05433522875"),
    ("یزد", "اصغر باغبان", "03537234960"),
    ("گناباد", "محمد نیکخواه", "05157253617"),
    ("گرگان", "عیسی بزی‌مقدم", "01732642990"),
    ("اردبیل", "ایرج خضری", "04533816336"),
    ("قزوین", "سعید دباغ‌کار", "02833292571"),
    ("همدان", "جلال بشیری", "08134383335"),
    ("انبار مرکزی تهران", "یاسر ربیعه", "02144543248"),
    ("دفتر غرب تهران", "سعید نجفی", "02144909605"),
    ("دفتر شرق تهران", "محمد صوفی", "02176710835"),
    ("دفتر کرج", "جواد بابامرادی", "02632705856"),
    ("دفتر ورامین", "محمد محمدی جعفری", "02136909225"),
    ("دفتر قم", "امین معصومی", "02532404984"),
    ("دفتر تخصصی تهران", "بهنام حفیظی", "02144542920"),
    ("دفتر خرم‌آباد", "علیرضا اسدی", "06633404288"),
    ("دفتر بندرعباس", "محمد معاذاللهی", "07632211280"),
    ("دفتر بوشهر", "آرش ترابیان", "07733320474"),
]

FERDOWS_HQ = [
    ("پخش فردوس", "مجید زجاجی — مدیرعامل (ferdowsdco)", "02166012695"),
    ("پخش فردوس", "حسین کوچک‌منش — قائم‌مقام / معاون مالی اداری (ferdowsdco)", ""),
]

# Mashhad FDO PDF rows: (company, manager, mgr_mobile, technical, tech_mobile, land)
# Parsed from v-drug.mums.ac.ir mashhad list
MASHHAD = [
    ("پخش اکسیر — مشهد", "حمیدرضا بابازاده", "09120236338", "پری‌ناز تبریزی", "09153101746", "05136220166"),
    ("پخش البرز — مشهد", "الهه امیری", "09155187584", "علیرضا فقیه", "09153018470", "05136667832"),
    ("الیت دارو — مشهد", "محمد صیامی", "09154309204", "سیدرضا اتحادیه", "09155138509", "05136147370"),
    ("اهورا پخش امین — مشهد", "کاوه", "09368097933", "آزاده زائرزاده", "09155212725", "05136513563"),
    ("آدورا طب — مشهد", "عادل", "09158010818", "علیرضا آراسته", "09155117228", "05135411974"),
    ("آوا دارو تدبیر شرق — مشهد", "مراتی", "09157894164", "بهزاد اسماعیلی", "09155042714", ""),
    ("آوا دارو تدبیر شرق — مشهد ۲", "مراتی", "09157894164", "فریبا خدایی", "09151020177", ""),
    ("به پخش داروی بهشهر — مشهد", "افتخاری", "09155151189", "مریم متولی حقیقی", "09153070950", "05133920223"),
    ("بهداشت سلامت آوندفر — مشهد", "ایمن", "09151115045", "زهرا ابوترابی", "09151105398", "05136095268"),
    ("بهرسان دارو — مشهد", "مجید خالقی", "09128508656", "فائزه تاج‌فرد", "09150692685", "05136515140"),
    ("بهستان پخش — مشهد", "وحیدرضا بهاری", "09151159140", "پدرام رحمانی", "09155129031", "05136512509"),
    ("پورا پخش — مشهد", "هادی وفادار", "09151117211", "مهشید وجدانی", "09153218010", "05136583320"),
    ("دارو بهداشت شفا آراد — مشهد", "سپیده فرودی", "09151018376", "جلال میرشاهی", "09153048849", "05136511087"),
    ("داروگستر ارمغان طبیعت — مشهد", "مهنوش ژیانی", "09155050475", "فرزین زیوری", "09157356051", "05136775498"),
    ("داروپخش — مشهد", "سیدعلی جعفری‌نژاد", "09155018189", "منیره حبیبی", "09153170132", "05133924871"),
    ("داروسازان التیام — مشهد", "محمد اسدی", "09153340162", "محسن محمودزاده", "09155012941", "05136237112"),
    ("داروگستر باریج اسانس — مشهد", "ارباب‌زاده", "09188634521", "هانیه پیروزمند", "09153004617", "05136676705"),
    ("داروگستر نخبگان — مشهد", "جلال ساری‌جعفری", "09153126606", "مونا زمانی", "09153170954", "05136659510"),
]


def main() -> None:
    hits: list[dict] = []
    for city, pe, tel in FERDOWS:
        hits.append(
            {
                "company": f"پخش فردوس — {city}",
                "person": f"{pe} — مدیر مرکز (ferdowsdco)",
                "mobile": None,
                "land": land(tel),
                "src": "ferdowds_warehouses",
            }
        )
    for co, pe, tel in FERDOWS_HQ:
        hits.append(
            {
                "company": co,
                "person": pe,
                "mobile": None,
                "land": land(tel) if tel else None,
                "src": "ferdowds_hq",
            }
        )
    for co, mgr, mm, tech, tm, tel in MASHHAD:
        hits.append(
            {
                "company": co,
                "person": f"{mgr} — مدیر شرکت (FDO مشهد)",
                "mobile": to_mobile(mm),
                "land": land(tel),
                "src": "mashhad_fdo_pdf",
            }
        )
        hits.append(
            {
                "company": co,
                "person": f"{tech} — مسئول فنی (FDO مشهد)",
                "mobile": to_mobile(tm),
                "land": land(tel),
                "src": "mashhad_fdo_pdf",
            }
        )

    existing = []
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if line.strip():
                existing.append(json.loads(line))
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
