# -*- coding: utf-8 -*-
"""Import Darogostar Tooba + Daroupakhsh Vet branches into person_id_hits.jsonl."""
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


def fetch(url: str, timeout: int = 30) -> str:
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")


def html_to_text(html: str) -> str:
    text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)
    return text.translate(DIGITS)


# Curated from live search snippets + page parse (dgtooba.com)
TOOBA = [
    ("دفتر مرکزی جهرم", "احمد کشافی", None, "07154335193"),
    ("شیراز", "سجاد زارع", "09173398101", "07136259456"),
    ("گیلان", "محمدرضا شریفی", "09120645254", None),
    ("مازندران", "حجت رضایی", "09111194970", "01132074099"),
    ("کهگیلویه و بویراحمد", "خداکرم کردستانی", "09173226457", "07433315079"),
    ("بوشهر", "روح الله شفیعی", "09164220259", "07791570517"),
    ("خراسان رضوی", "حمیرا همائی شاد", "09156953815", "05136651317"),
    ("لرستان", "حسین رضایی ارشد", "09395567117", None),
    ("اصفهان", "محمد امین نوذریان", "09133197981", "03191092132"),
    ("خوزستان", "محمد نیکوسیر", "09171900135", "06133908663"),
    ("هرمزگان", "کتایون آریانژاد", "09375223117", "07632560204"),
    ("کرمانشاه", "مهدی رمضانی", "09184238981", None),
    ("زاهدان", "آرزو شه بخش", "09333130306", None),
    ("تهران", "امید قنبری", "09177914959", None),
    ("آذربایجان شرقی", "مجید روغن فروش ممقانی", "09141028349", None),
    ("کرمان", "حسین توفیقی", "09171087664", "03432719562"),
]

# Daroupakhsh Vet (ddvetco.com) branch heads — mostly land + person
DDVET = [
    ("تبریز", "دکتر امیررضا نارونی", None, None),
    ("اردبیل", "دکتر نیما عالمی سرج", None, "04533644030"),
    ("اصفهان", "دکتر حامد هدایتی قره‌چه", "09132260159", None),
    ("کرج سرخه‌حصار", "محسن محمودی", None, "02146852185"),
    ("ایلام", "احسان جوکار", None, "08433310262"),
    ("تهران/کرج", "دکتر عطا اله زندی", None, "02146852185"),
    ("شهرکرد", "دکتر امیر حاتمی بلداجی", None, "03832253283"),
    ("مشهد", "مهندس محمد صادقی دخت", None, "05136147329"),
    ("زنجان", "فریدون سلیمی ویری", None, "02433452499"),
    ("شیراز", "دکتر احسان حقیقی فرجام", None, "07137740122"),
    ("قزوین", "دکتر عرفان بیگوند", None, "02833575945"),
    ("قم", "دکتر کیمیا وطن‌خواه", None, "02536674384"),
    ("سنندج", "سیده افسانه محمودی", None, None),
]


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


def main():
    hits: list[dict] = []

    # Try refresh Tooba page for any extra mobiles
    try:
        html = fetch("https://dgtooba.com/")
        (RAW / "dgtooba.html").write_text(html, encoding="utf-8")
        text = html_to_text(html)
        (RAW / "dgtooba_text.txt").write_text(text, encoding="utf-8")
        # opportunistic parse: "مدیر ... NAME" near mobile
        for m in re.finditer(
            r"مدیر(?:\s*شعبه|\s*عامل)?\s*[:：]?\s*([آ-یA-Za-z\u0600-\u06FF\s]{4,45}).{0,80}?((?:0?9\d{9}))",
            text,
            re.S,
        ):
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
            mb = to_mobile(m.group(2))
            if pe and mb:
                hits.append(
                    {
                        "company": "داروگستر طوبی — شعبه (وب)",
                        "person": f"{pe} (مدیر شعبه داروگستر طوبی)",
                        "mobile": mb,
                        "land": None,
                        "src": "dgtooba_parse",
                    }
                )
    except Exception as e:
        print({"dgtooba_fetch_error": str(e)})

    for city, person, mob, ld in TOOBA:
        hits.append(
            {
                "company": f"داروگستر طوبی — شعبه {city}",
                "person": f"{person} (مدیر شعبه داروگستر طوبی {city})",
                "mobile": to_mobile(mob),
                "land": land(ld),
                "src": "dgtooba_curated",
            }
        )

    # Fetch ddvet page
    try:
        html = fetch("https://ddvetco.com/%D8%B4%D8%B9%D8%A8-%D9%88-%D9%86%D9%85%D8%A7%DB%8C%D9%86%D8%AF%DA%AF%DB%8C%D9%87%D8%A7/")
        (RAW / "ddvet_branches.html").write_text(html, encoding="utf-8")
        text = html_to_text(html)
        (RAW / "ddvet_branches_text.txt").write_text(text, encoding="utf-8")
        # parse head + phone pairs
        for m in re.finditer(
            r"(?:رئیس|سرپرست)\s*شعبه\s*[:：]?\s*([آ-یA-Za-z\u0600-\u06FF\s‌]{4,50}).{0,200}?(?:شماره تماس|تلفن)?\s*[:：]?\s*((?:0?\d{2,3}[\s\-]?\d{6,8})|(?:0?9\d{9}))",
            text,
            re.S,
        ):
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
            phone = m.group(2)
            mb = to_mobile(phone)
            ld = None if mb else land(phone)
            if pe and (mb or ld):
                hits.append(
                    {
                        "company": "توزیع داروهای دامی داروپخش — شعبه",
                        "person": f"{pe} (رئیس شعبه داروپخش دامی)",
                        "mobile": mb,
                        "land": ld,
                        "src": "ddvet_parse",
                    }
                )
    except Exception as e:
        print({"ddvet_fetch_error": str(e)})

    for city, person, mob, ld in DDVET:
        hits.append(
            {
                "company": f"توزیع داروهای دامی داروپخش — شعبه {city}",
                "person": f"{person} (رئیس شعبه داروپخش دامی {city})",
                "mobile": to_mobile(mob),
                "land": land(ld),
                "src": "ddvet_curated",
            }
        )

    # Adorateb sample managers (person+land) — upgrade co_pe
    adorateb = [
        ("گرگان", "حسین محمدی", None, "01732470181"),
        ("تبریز", "امید اناری", None, "04136377037"),
    ]
    for city, person, mob, ld in adorateb:
        hits.append(
            {
                "company": f"آدوراطب — مرکز {city}",
                "person": f"{person} (مدیر مرکز آدوراطب {city})",
                "mobile": to_mobile(mob),
                "land": land(ld),
                "src": "adorateb_web",
            }
        )

    # Razi HQ executives (person + land only — still co_pe KPI)
    razi_exec = [
        ("حمید بیرانوند", "مدیرعامل پخش رازی", "02166706812"),
        ("عرفان فراهانی", "معاون سرمایه انسانی و پشتیبانی پخش رازی", "02166704105"),
        ("لطف‌الله رشیدی", "معاون فروش و بازاریابی پخش رازی", "02166710891"),
        ("مهتاب محسنی", "معاون بازرگانی پخش رازی", "02166750463"),
        ("سعید علیجانی", "معاون مالی و اقتصادی پخش رازی", "02166704105"),
        ("حسین عاشوری", "سرپرست معاونت طرح و توسعه پخش رازی", "02166726765"),
        ("هوشنگ غریب پور", "معاون اجرایی پخش رازی", None),
        ("لاله علیوند", "مسئول فنی پخش رازی", "02166706840"),
        ("عباس اخلاقی", "مدیر IT پخش رازی", "02166726764"),
        ("ابراهیم وارثی پور", "مدیر روابط عمومی پخش رازی", "02163457000"),
    ]
    for person, role, ld in razi_exec:
        hits.append(
            {
                "company": "پخش رازی — دفتر مرکزی",
                "person": f"{person} ({role})",
                "mobile": None,
                "land": land(ld),
                "src": "razico_contact",
            }
        )

    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
