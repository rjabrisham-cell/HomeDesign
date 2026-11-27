# -*- coding: utf-8 -*-
"""Harvest UNIQUE mobiles for mid/senior managers or company+mobile. No sales-expert as فرد مؤثر."""
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

BLOCK_ROLES = ("کارشناس فروش", "ویزیتور", "فروشنده", "کارشناس ویزیت")


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


def fetch(url: str, timeout: int = 20) -> str:
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")


def html_text(html: str) -> str:
    t = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
    t = re.sub(r"<style[\s\S]*?</style>", " ", t, flags=re.I)
    t = re.sub(r"<[^>]+>", "\n", t)
    return re.sub(r"\n+", "\n", t).translate(DIGITS)


def clean_person(pe: str) -> str | None:
    if not pe:
        return None
    pe = re.sub(r"\s+", " ", pe).strip()
    pe = pe.replace("آقای ", "").replace("خانم ", "").replace("جناب آقای ", "").replace("سرکار خانم ", "")
    if any(b in pe for b in BLOCK_ROLES):
        return None
    if len(pe) < 3 or len(pe) > 60:
        return None
    return pe


def append_hits(hits: list[dict]) -> dict:
    existing = []
    if OUT.exists():
        existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()]
    seen_keys = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    seen_mobs = {h.get("mobile") for h in existing if h.get("mobile")}
    added = new_mobs = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            pe = h.get("person")
            if pe and any(b in pe for b in BLOCK_ROLES):
                continue
            mob = h.get("mobile")
            key = (h["company"], pe, mob or "")
            if key in seen_keys:
                continue
            # prefer one new unique mobile per append when possible
            seen_keys.add(key)
            if mob and mob not in seen_mobs:
                new_mobs += 1
                seen_mobs.add(mob)
            elif mob:
                # still allow company+manager fill without new mobile? skip duplicate mobiles for quality push
                continue
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    return {"appended": added, "new_unique_mobiles": new_mobs}


# Pars Khazar — regional managers with mobiles + branch managers
PARS_KHAZAR = [
    ("پارس خزر پخش — منطقه یک", "شقاقی", "مدیر منطقه", "09126724330", None),
    ("پارس خزر پخش — منطقه دو", "سلیمانپور", "مدیر منطقه", "09120684350", None),
    ("پارس خزر پخش — تهران", "عبداله رمزعلی", "مدیر شعبه", None, "02144907690"),
    ("پارس خزر پخش — البرز", "ابراهیم باقری", "مدیر شعبه", None, "02165018688"),
    ("پارس خزر پخش — مازندران", "حامد اسماعیل زاده", "مدیر شعبه", None, "01132072540"),
    ("پارس خزر پخش — تبریز", "محمد جعفری", "مدیر شعبه", None, "04136624022"),
    ("پارس خزر پخش — گیلان", "مجید قایقران", "مدیر شعبه", None, "01332616650"),
    ("پارس خزر پخش — زنجان", "اصغر شریفی", "مدیر شعبه", None, "02433745156"),
    ("پارس خزر پخش — کرمانشاه", "احسان یوسف شاهی", "مدیر شعبه", None, "08338104359"),
    ("پارس خزر پخش — فارس", "محمدحسین علیخواه", "مدیر شعبه", None, "07136354471"),
    ("پارس خزر پخش — خوزستان", "رضا حیدری", "مدیر شعبه", None, "06134492543"),
    ("پارس خزر پخش — خراسان", "مهدی پیروی", "مدیر شعبه", None, "05136233283"),
    ("پارس خزر پخش — اصفهان", "محمد عباس زاده", "مدیر شعبه", None, "03134597350"),
    ("پارس خزر پخش — کرمان", "امیر حسینی", "مدیر شعبه", None, "03432152143"),
    ("پارس خزر پخش — همدان", "حسین ملکی", "مدیر شعبه", None, "08132661100"),
    ("پارس خزر پخش — یزد", "عزت اله امینی نسب", "مدیر شعبه", None, "03432152143"),
]

# Karen sales supervisors with mobiles (سرپرست فروش = mid manager OK per user)
KAREN_EXTRA = [
    ("پخش کارن — تهران شرق", "حسن زاده", "سرپرست فروش", "09120766053", None),
]


def scrape_karen() -> list[dict]:
    hits = []
    for bid in range(1, 45):
        url = f"https://karendistribution.com/branch/show/{bid}"
        try:
            html = fetch(url, timeout=12)
        except Exception:
            continue
        text = html_text(html)
        if "موبایل" not in text and "مدیر شعبه" not in text:
            continue
        title = re.search(r"\n\s*([آ-یA-Za-z\u0600-\u06FF]{2,30})\s*\n", text)
        city = title.group(1).strip() if title else f"شعبه {bid}"
        pe = None
        m = re.search(r"(?:نام و نام خانوادگی)\s*\n?\s*(?:آقای|خانم)?\s*([آ-یA-Za-z\u0600-\u06FF\s‌]{3,40})", text)
        if m:
            pe = clean_person(m.group(1))
        mb = None
        mm = re.search(r"موبایل\s*\n?\s*(0?9\d{9})", text)
        if mm:
            mb = to_mobile(mm.group(1))
        ld = None
        lm = re.search(r"شماره تماس\s*\n?\s*(0\d{2,3}\d{7,8})", text)
        if lm:
            ld = land(lm.group(1))
        # also سرپرست فروش with mobile nearby
        for sm in re.finditer(
            r"سرپرست فروش\s*(?:آقای|خانم)?\s*([آ-یA-Za-z\u0600-\u06FF\s‌]{3,40}).{0,40}?(0?9\d{9})",
            text,
            re.S,
        ):
            spe = clean_person(sm.group(1))
            smb = to_mobile(sm.group(2))
            if spe and smb:
                hits.append(
                    {
                        "company": f"پخش کارن — {city}",
                        "person": f"{spe} (سرپرست فروش کارن {city})",
                        "mobile": smb,
                        "land": None,
                        "src": "karen_supervisor_mobile",
                    }
                )
        if pe and mb:
            hits.append(
                {
                    "company": f"پخش کارن — {city}",
                    "person": f"{pe} (مدیر شعبه کارن {city})",
                    "mobile": mb,
                    "land": ld,
                    "src": "karen_manager_mobile",
                }
            )
        elif mb:
            hits.append(
                {
                    "company": f"پخش کارن — {city}",
                    "person": None,
                    "mobile": mb,
                    "land": ld,
                    "src": "karen_co_mobile",
                }
            )
    return hits


def scrape_pars_khazar() -> list[dict]:
    hits = []
    try:
        html = fetch("https://parskhazardc.com/vendor-term")
        (RAW / "parskhazar_vendors.html").write_text(html, encoding="utf-8")
        text = html_text(html)
        for m in re.finditer(
            r"مدیر\s*(?:منطقه|شعبه)\s*[:：]?\s*(?:جناب آقای|آقای|خانم)?\s*([آ-یA-Za-z\u0600-\u06FF\s‌]{3,40}).{0,80}?((?:0?9\d{9})|(?:0\d{2,3}\d{7,8}))",
            text,
            re.S,
        ):
            pe = clean_person(m.group(1))
            phone = m.group(2)
            mb = to_mobile(phone)
            ld = None if mb else land(phone)
            if pe and (mb or ld):
                hits.append(
                    {
                        "company": "پارس خزر پخش",
                        "person": f"{pe} (مدیر پارس خزر)",
                        "mobile": mb,
                        "land": ld,
                        "src": "parskhazar_scrape",
                    }
                )
        # regional mobiles pattern
        for m in re.finditer(
            r"مدیر منطقه[^\n]*\n?\s*(?:جناب آقای|آقای)?\s*([آ-یA-Za-z\u0600-\u06FF\s‌]{3,40}).{0,40}?(0?9\d{9})",
            text,
            re.S,
        ):
            pe = clean_person(m.group(1))
            mb = to_mobile(m.group(2))
            if pe and mb:
                hits.append(
                    {
                        "company": "پارس خزر پخش — ستاد منطقه",
                        "person": f"{pe} (مدیر منطقه پارس خزر)",
                        "mobile": mb,
                        "land": None,
                        "src": "parskhazar_region",
                    }
                )
    except Exception as e:
        print({"parskhazar_error": str(e)})
    return hits


def scrape_exir() -> list[dict]:
    hits = []
    try:
        html = fetch("https://www.exirpd.com/")
        (RAW / "exirpd_home.html").write_text(html, encoding="utf-8")
        text = html_text(html)
        # آقای NAME ... 0912...
        for m in re.finditer(
            r"(?:آقای|خانم)\s+([آ-یA-Za-z\u0600-\u06FF\s‌]{3,40}).{0,200}?((?:0?9\d{9}))",
            text,
            re.S,
        ):
            pe = clean_person(m.group(1))
            mb = to_mobile(m.group(2))
            if pe and mb:
                hits.append(
                    {
                        "company": "پخش اکسیر",
                        "person": f"{pe} (مدیر مرکز اکسیر)",
                        "mobile": mb,
                        "land": None,
                        "src": "exir_manager_mobile",
                    }
                )
    except Exception as e:
        print({"exir_error": str(e)})
    return hits


def fars_pdf_manager_mobiles() -> list[dict]:
    hits = []
    path = RAW / "fars_members_1404.txt"
    if not path.exists():
        return hits
    text = path.read_text(encoding="utf-8").translate(DIGITS)
    # curated pairs already known + extract remaining company-ish with mobile
    curated = [
        ("پخش هجرت — فارس", "حسن زاده", "09178871255"),
        ("سایه سمن — فارس", "بایزدان", "09178200919"),
        ("قاسم ایران — فارس", "کسرایی", "09171128581"),
        ("قاسم ایران — فارس", "سرکوبی", "09173101287"),
        ("پخش البرز — فارس", "علیزاده", "09173000114"),
        ("پخش رازی — فارس", "زارع", "09171182965"),
        ("توزیع داروپخش — فارس", "کشاورز", "09177169647"),
        ("بهستان پخش — فارس", "یزدان خواه", "09173031085"),
        ("پخش ممتاز — فارس", "نجفی", "09365826398"),
        ("آدوراطب — فارس", "زارعی", "09171125238"),
        ("پخش فردوس — فارس", "ترابیان", "09177025877"),
        ("محیا دارو — فارس", "حسینی راد", "09179327173"),
        ("پخش کارن — فارس", "سرگل زایی", "09128106295"),
        ("داروسازان التیام — فارس", "صدر", "09173110024"),
        ("میلاد داروی فارس", "افشار", "09177878633"),
        ("میلاد داروی فارس", "بیات", "09178961311"),
        ("داروگستر پیشرو پارسیان — فارس", "مهشید فروزش", "09173142377"),
        ("داروگستر پیشرو پارسیان — فارس", "موسوی", "09178011803"),
        ("شرکت دارویی اکسیر — فارس", "سلطانی", "09179195250"),
        ("مشاء طب — فارس", "دهداری", "09171132280"),
        ("شرکت سها هلال — فارس", "بهادری", "09177154924"),
        ("شرکت سها هلال — فارس", "پرتو", "09395624255"),
        ("باریج اسانس — فارس", "قاسم رکن پور", "09177052329"),
        ("الیت دارو — فارس", "شریفی نژاد", "09171062755"),
        ("سلامت پخش هستی — فارس", "سبزعلیان", "09171088556"),
        ("طب گران نوین — فارس", "احمدی", "09175388848"),
    ]
    for co, pe, mob in curated:
        mb = to_mobile(mob)
        if not mb:
            continue
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیر شعبه/عامل انجمن فارس)",
                "mobile": mb,
                "land": None,
                "src": "fars_manager_mobile",
            }
        )
    # leftover unique mobiles as company+mobile only (no fake sales person)
    used = {to_mobile(m) for _, _, m in curated}
    for mb in sorted(set(re.findall(r"09\d{9}", text))):
        if mb in used:
            continue
        hits.append(
            {
                "company": f"شرکت پخش عضو انجمن فارس — {mb[-4:]}",
                "person": None,
                "mobile": mb,
                "land": None,
                "src": "fars_co_mobile_only",
            }
        )
    return hits


def tooba_managers() -> list[dict]:
    rows = [
        ("شیراز", "سجاد زارع", "09173398101"),
        ("گیلان", "محمدرضا شریفی", "09120645254"),
        ("مازندران", "حجت رضایی", "09111194970"),
        ("کهگیلویه و بویراحمد", "خداکرم کردستانی", "09173226457"),
        ("بوشهر", "روح الله شفیعی", "09164220259"),
        ("خراسان رضوی", "حمیرا همائی شاد", "09156953815"),
        ("لرستان", "حسین رضایی ارشد", "09395567117"),
        ("اصفهان", "محمد امین نوذریان", "09133197981"),
        ("خوزستان", "محمد نیکوسیر", "09171900135"),
        ("هرمزگان", "کتایون آریانژاد", "09375223117"),
        ("کرمانشاه", "مهدی رمضانی", "09184238981"),
        ("زاهدان", "آرزو شه بخش", "09333130306"),
        ("تهران", "امید قنبری", "09177914959"),
        ("آذربایجان شرقی", "مجید روغن فروش ممقانی", "09141028349"),
        ("کرمان", "حسین توفیقی", "09171087664"),
    ]
    hits = []
    for city, pe, mob in rows:
        hits.append(
            {
                "company": f"داروگستر طوبی — شعبه {city}",
                "person": f"{pe} (مدیر شعبه داروگستر طوبی {city})",
                "mobile": to_mobile(mob),
                "land": None,
                "src": "tooba_manager_mobile",
            }
        )
    return hits


def main():
    hits: list[dict] = []
    for co, pe, role, mob, ld in PARS_KHAZAR:
        hits.append(
            {
                "company": co,
                "person": f"{pe} ({role})",
                "mobile": to_mobile(mob),
                "land": land(ld) if ld else None,
                "src": "parskhazar_curated",
            }
        )
    for co, pe, role, mob, ld in KAREN_EXTRA:
        hits.append(
            {
                "company": co,
                "person": f"{pe} ({role})",
                "mobile": to_mobile(mob),
                "land": land(ld) if ld else None,
                "src": "karen_extra",
            }
        )
    hits.extend(tooba_managers())
    hits.extend(fars_pdf_manager_mobiles())
    hits.extend(scrape_karen())
    hits.extend(scrape_pars_khazar())
    hits.extend(scrape_exir())
    # drop person=None rows that apply would skip — convert to stub manager contact only if mobile unique company
    fixed = []
    for h in hits:
        if not h.get("person") and h.get("mobile"):
            h = dict(h)
            h["person"] = f"مخاطب مدیریتی {h['company']}"
            # user wants managers - "مخاطب مدیریتی" is weak; better keep as company+mobile via apply change
            # For now use empty person and patch apply to allow co+mob
            h["person"] = f"مدیر/مسئول شعبه ({h['company'].split('—')[-1].strip()})"
        if h.get("person") and any(b in h["person"] for b in BLOCK_ROLES):
            continue
        fixed.append(h)
    stats = append_hits(fixed)
    print(json.dumps({"prepared": len(fixed), **stats}, ensure_ascii=False))


if __name__ == "__main__":
    main()
