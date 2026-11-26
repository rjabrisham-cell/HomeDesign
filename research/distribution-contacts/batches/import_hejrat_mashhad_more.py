# -*- coding: utf-8 -*-
"""Scrape Hejrat centers + remaining Mashhad FDO rows."""
from __future__ import annotations

import json
import re
import urllib.request
from pathlib import Path

OUT = Path(r"D:\HomeDesign\research\distribution-contacts\batches\person_id_hits.jsonl")
RAW = Path(r"D:\HomeDesign\research\distribution-contacts\raw")
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
        elif len(d) == 8:
            d = "021" + d
    return d if d.startswith("0") and 10 <= len(d) <= 11 else None


# Additional Mashhad triples from PDF remainder
MASHHAD_MORE = [
    ("بهشاد آراءشرق مدهوشی — مشهد", "مدهوشی", "09151125171", "رضا ژیان باقرپور", "09358600110", "05136147260"),
    ("پونل برسام — مشهد", "نوید عالیی", "09121329810", "نگار ذوالجلالی", "09151741121", "05138830798"),
    ("تجهیزگران طب نوین — مشهد", "محمدیان", "09153119368", "فاضل حسینی", "09155206684", "05136676442"),
    ("توسعه تجارت جلال آرا — مشهد", "صباحی", "09152005416", "مریم قربانی", "09151191245", "05132403575"),
    ("خدمات دارویی رضوی — مشهد", "مواجی", "09155064018", "ذبیح‌اله روحی", "09153173970", "05136511320"),
    ("داروگستر پارس کیمیا — مشهد", "مهدی چراغی", "09153134935", "مازیار صبحیان", "09156248985", "05136576174"),
    ("داروگستر یاسین — مشهد", "سلیمانی", "09151063006", "سیما کوهستانی", "09151710180", "05136517502"),
    ("دایا دارو — مشهد", "طیبی", "09153096119", "سمیرا غدیرزاده", "09155190480", "05136676462"),
    ("دی دارو امید — مشهد", "دهقان‌نژاد", "09362934547", "مهدیه مرکبی", "09151134869", ""),
    ("دینه دارو درمان — مشهد", "عین‌افشار", "09155096366", "انسیه نوروزی", "09153189084", "05136652372"),
    ("پخش رازی — مشهد", "هاشمی", "09126722206", "فریبا آزادی", "09151246629", "05133928814"),
    ("سایه سمن — مشهد", "شمع‌ریز", "09155036254", "کوثری", "09151513025", "05135422160"),
    ("سبز دارو سپاهان — مشهد", "تازیکی", "09154030035", "آرزو سبحانی", "09155584785", ""),
    ("سلامت پخش هستی — مشهد", "زارع", "09155093358", "اطلس روحی", "09151011782", "05136512241"),
    ("سلامت گستر شایان اعتماد — مشهد", "طالبی", "09158299919", "مطهره صباغ", "09397149525", ""),
    ("سها هلال احمر — مشهد", "دادخدازاده", "09151135403", "فاطمه شهریاری", "09120915410", "05136147295"),
    ("سینا پخش ژن — مشهد", "رویین‌تن", "09361793382", "سارا عامری", "09151371051", "05136779552"),
    ("سینا پخش صبا — مشهد", "جعفریان", "09151779070", "روح‌ا. عسگری", "09153388834", "05136513370"),
    ("شفا پخش هیرکان — مشهد", "تازیکی", "09154030035", "رویا جمشیدی", "09391635850", ""),
    ("طبیعت زنده (سینره) — مشهد", "عزیزی", "09120141917", "شکوفه لوایی", "09151576882", ""),
    ("فراگیر پخش دماوند — مشهد", "پرویز غرباغی", "09151571476", "متولی حقی", "", "05135022780"),
    ("پخش فردوس — مشهد", "وحید سجادی", "09153136508", "نصرت‌اله دانشمند", "", "05133929049"),
    ("قاسم ایران — مشهد", "اطهم خانی", "09155006442", "داود ممقانی ثانی", "09155577162", "05136578434"),
    ("پخش کارن — مشهد", "حسینی", "09158049099", "الهه غلامی", "09155981720", "05136662021"),
    ("گیتی سلامت آریا — مشهد", "علیرضا عطایی", "09195760206", "احسان جلیلی", "09158222848", "05136512189"),
    ("لیوار — مشهد", "امیدرضا آذین‌فرد", "09124392026", "سمیه بیدار", "", "05136513070"),
    ("محیا دارو — مشهد", "حسن‌زاده", "09155115555", "مرضیه قره‌چاهی", "", "05136582654"),
    ("مشاء طب طهران — مشهد", "مصطفی ظریف مجیدی", "09154050916", "ناصحی", "", "05136583320"),
    ("مشکات فارمد توزیع — مشهد", "محمدی‌پور", "09151010709", "علی تمیزی", "09153070386", ""),
    ("پخش ممتاز — مشهد", "پاشایی", "09121083479", "سجادی", "09151105365", "05136663844"),
    ("ایده گستران زمان — مشهد", "رامین آخوندی", "09155253410", "احمد باوفا", "09153130961", "05134513963"),
    ("مفید دارو آریا فارمد — مشهد", "اصیلی", "09128504535", "سارا حسینی", "09151579513", ""),
    ("مهبان دارو — مشهد", "محمدرضا موذنی", "09151193772", "المیرا محتشمی", "09154154813", "05136512182"),
    ("میردامادی و شرکاء — مشهد", "مردان", "09151158835", "مونا رضازاده", "09159008158", "05136666554"),
    ("نیک فردای سالم (نفس) — مشهد", "احمدی", "09151176356", "مرضیه مظلوم‌زاده", "09105090790", "05136513612"),
    ("ویهان طب پارس — مشهد", "رضا دهقان‌نژاد", "09151134869", "سمیرا صدیقی", "09159175982", "05137641020"),
    ("پخش هجرت — مشهد", "رئوف", "09155056725", "حسن نامجو", "09151308030", "05133921214"),
    ("هنار دارو — مشهد", "محمود جعفری‌نژاد", "09151674540", "حسن‌زاده", "09151681315", "05136696030"),
    ("یارا طب ثامن — مشهد", "ضیائی", "09151096424", "شهلا پرند", "09123387642", "05135424267"),
    ("یاقوت سلامت — مشهد", "فرحزاد", "09155005157", "میترا لعل‌محمدی", "09155116049", ""),
]


def scrape_hejrat() -> list[dict]:
    hits = []
    for i in range(1, 40):
        url = f"https://hejratco.com/DistributeCenter/Details/{i}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=15) as r:
                html = r.read().decode("utf-8", "ignore")
        except Exception:
            continue
        # manager name near "نام مدیر" or heading
        pe = None
        m = re.search(r"اطلاعات مدیر مرکز\s*([^<\n]{0,40})", html)
        # look for h5 after manager section
        m2 = re.search(
            r"اطلاعات مدیر مرکز[\s\S]{0,400}?<h\d[^>]*>\s*([^<]{3,60})\s*</h",
            html,
            re.I,
        )
        if m2:
            pe = re.sub(r"\s+", " ", m2.group(1)).strip()
        if not pe:
            m3 = re.search(r"نام مدیر\s*[:：]?\s*</[^>]+>\s*<[^>]+>\s*([^<]{3,60})", html)
            if m3:
                pe = re.sub(r"\s+", " ", m3.group(1)).strip()
        # city from title or breadcrumb
        title_m = re.search(r"<title>([^<]+)</title>", html, re.I)
        city = None
        # look for center name
        cm = re.search(r"مرکز\s+([آ-یA-Za-z\s]{2,30})", html)
        if cm:
            city = cm.group(1).strip()
        if not city and title_m:
            city = re.sub(r"\s+", " ", title_m.group(1)).strip()[:40]
        # phones
        lands = []
        for mm in re.finditer(r"(?<!\d)(0\d{2}\d{7,8})(?!\d)", html.translate(DIGITS)):
            d = mm.group(1)
            if d.startswith("09"):
                continue
            if d not in lands:
                lands.append(d)
        if pe and pe not in ("نام مدیر", "-", "اطلاعات تماس"):
            hits.append(
                {
                    "company": f"پخش هجرت — مرکز {city or i}",
                    "person": f"{pe} — مدیر مرکز (hejratco)",
                    "mobile": None,
                    "land": lands[0] if lands else None,
                    "src": "hejrat_centers",
                }
            )
    (RAW / "hejrat_centers.json").write_text(
        json.dumps(hits, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return hits


def main() -> None:
    hits: list[dict] = []
    for co, mgr, mm, tech, tm, tel in MASHHAD_MORE:
        hits.append(
            {
                "company": co,
                "person": f"{mgr} — مدیر شرکت (FDO مشهد)",
                "mobile": to_mobile(mm),
                "land": land(tel),
                "src": "mashhad_fdo_more",
            }
        )
        if tech:
            hits.append(
                {
                    "company": co,
                    "person": f"{tech} — مسئول فنی (FDO مشهد)",
                    "mobile": to_mobile(tm),
                    "land": land(tel),
                    "src": "mashhad_fdo_more",
                }
            )
    hej = scrape_hejrat()
    hits.extend(hej)

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
    print(
        json.dumps(
            {"batch": len(hits), "appended": added, "with_mobile": with_mob, "hejrat": len(hej)},
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
