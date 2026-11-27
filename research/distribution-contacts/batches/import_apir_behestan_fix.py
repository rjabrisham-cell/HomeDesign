# -*- coding: utf-8 -*-
"""Fix APIR board parse (th/td) + scrape Behestan branch managers + HQ team."""
from __future__ import annotations

import json
import re
import urllib.request
from pathlib import Path
from urllib.parse import unquote

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


def fetch(url: str, timeout: int = 20) -> str:
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")


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


def scrape_apir() -> list[dict]:
    hits = []
    for n in range(1, 41):
        url = f"https://apir.ir/branches/branch-{n}"
        try:
            html = fetch(url, timeout=15)
        except Exception:
            continue
        if "scope=\"row\"" not in html and "هیئت مدیره" not in html:
            continue
        # province
        pm = re.search(r"<h1[^>]*>\s*استان\s*([^<]+)</h1>", html, re.I)
        if not pm:
            pm = re.search(r"انجمن\s+([آ-یA-Za-z\u0600-\u06FF\s]{2,40})", html)
        province = re.sub(r"\s+", " ", pm.group(1)).strip() if pm else f"شعبه {n}"

        lands = []
        # footer/header phones like 02144009809
        for mm in re.finditer(r"(?<!\d)(0\d{2}\d{8})(?!\d)", html.translate(DIGITS)):
            d = mm.group(1)
            if not d.startswith("09") and d not in lands:
                lands.append(d)
        for mm in re.finditer(r"(?<!\d)(\d{5,8})\s*[-–]\s*(0?\d{2,3})(?!\d)", html.translate(DIGITS)):
            local, area = mm.group(1), mm.group(2)
            if not area.startswith("0"):
                area = "0" + area
            ld = land(area + local)
            if ld and ld not in lands:
                lands.append(ld)

        page = 0
        for m in re.finditer(
            r"<th[^>]*scope=\"row\"[^>]*>\s*(?:آقای|خانم|سرکار خانوم|سرکار خانم)?\s*([^<]{3,60})\s*</th>\s*<td[^>]*>\s*([^<]{2,80})\s*</td>",
            html,
            re.I,
        ):
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
            role = re.sub(r"\s+", " ", m.group(2)).strip()
            if pe == "نام و نام خانوادگی" or role == "سمت":
                continue
            pe = (
                pe.replace("آقای ", "")
                .replace("خانم ", "")
                .replace("سرکار خانوم ", "")
                .replace("سرکار خانم ", "")
            )
            hits.append(
                {
                    "company": f"انجمن ملی صنعت پخش — {province}",
                    "person": f"{pe} ({role} انجمن {province})",
                    "mobile": None,
                    "land": lands[0] if lands else None,
                    "src": "apir_board_th",
                    "url": url,
                }
            )
            page += 1
        if page:
            print({"apir": url, "hits": page, "province": province})
    return hits


BEHESTAN_STAFF = [
    ("دکتر محمدهادی کمالی", "مدیرعامل"),
    ("رضا اشجعی", "معاون فروش"),
    ("رئوف اسدیان", "مدیر خرید"),
    ("بابک میثاقی", "مدیر فناوری اطلاعات"),
    ("هومن آجودانی", "مدیر حسابداری فروش"),
    ("صبا ساسانی", "مدیر حسابداری مالی"),
    ("مهدی احمدپور علمداری", "مدیر فنی و پشتیبانی"),
    ("مهیار قربانی", "مدیر برنامه ریزی استراتژیک"),
    ("مجتبی ابوالفتحی", "مدیر حسابرسی داخلی"),
    ("لیدا بهادر", "مدیر فنی دارویی"),
    ("حسین پرندوار", "مدیر حقوقی"),
]


def scrape_behestan() -> list[dict]:
    hits = []
    for pe, role in BEHESTAN_STAFF:
        hits.append(
            {
                "company": "بهستان پخش — دفتر مرکزی",
                "person": f"{pe} ({role} بهستان پخش)",
                "mobile": None,
                "land": land("02184054000"),
                "src": "behestan_team",
            }
        )

    # discover branch URLs from sitemap / search pages via WP pages list heuristics
    seeds = [
        "https://behestanpakhsh.com/pages/",
        "https://behestanpakhsh.com/sitemap.xml",
        "https://behestanpakhsh.com/sitemap_index.xml",
        "https://behestanpakhsh.com/wp-sitemap-posts-page-1.xml",
        "https://behestanpakhsh.com/page-sitemap.xml",
    ]
    urls = set()
    for seed in seeds:
        try:
            html = fetch(seed, timeout=20)
            (RAW / ("behestan_" + re.sub(r"\W+", "_", seed)[-40:] + ".txt")).write_text(html[:200000], encoding="utf-8")
            for u in re.findall(r"https://behestanpakhsh\.com/pages/[^\s\"'<>]+", html):
                urls.add(u.rstrip("/"))
            for u in re.findall(r"/pages/[^\s\"'<>]+", html):
                if "pages" in u:
                    urls.add("https://behestanpakhsh.com" + u if u.startswith("/") else u)
        except Exception as e:
            print({"behestan_seed_error": seed, "err": str(e)})

    # curated known province pages if discovery weak
    curated = [
        "گیلان",
        "آذربایجان شرقی",
        "اصفهان بهستان",
        "کرمان",
        "فارس",
        "بوشهر",
        "همدان",
        "تهران",
        "مشهد",
        "اهواز",
        "کرج",
        "قم",
        "یزد",
        "کرمانشاه",
        "ارومیه",
        "ساری",
        "رشت",
        "زاهدان",
        "بندرعباس",
        "اردبیل",
        "زنجان",
        "قزوین",
        "لرستان",
        "گلستان",
        "ایلام",
        "شهرکرد",
        "سنندج",
        "بجنورد",
        "بیرجند",
        "سمنان",
        "یاسوج",
    ]
    # Also try WP REST
    try:
        api = fetch("https://behestanpakhsh.com/wp-json/wp/v2/pages?per_page=100", timeout=25)
        for item in json.loads(api):
            link = item.get("link") or ""
            title = re.sub(r"<[^>]+>", "", item.get("title", {}).get("rendered", ""))
            if any(k in title for k in ("مرکز پخش", "شعبه", "مرکز")) or "/pages/" in link:
                if link:
                    urls.add(link.rstrip("/"))
    except Exception as e:
        print({"behestan_api_error": str(e)})

    print({"behestan_urls": len(urls)})
    for url in sorted(urls):
        if "team" in url or "درباره" in unquote(url):
            continue
        try:
            html = fetch(url, timeout=15)
        except Exception:
            continue
        if "مدیر شعبه" not in html and "سرپرست فروش" not in html:
            continue
        text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
        text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
        text = re.sub(r"<[^>]+>", "\n", text)
        text = re.sub(r"\n+", "\n", text).translate(DIGITS)
        tm = re.search(r"(?:مرکز پخش دارو|شعبه)\s+([^\n]{2,40})", text)
        city = tm.group(1).strip() if tm else "نامشخص"
        lands = []
        for mm in re.finditer(r"(?<!\d)(0\d{2,3}\d{7,8})(?!\d)", text):
            d = mm.group(1)
            if d.startswith("0218405"):
                continue
            if not d.startswith("09") and d not in lands:
                lands.append(d)
        # people: #### Name \n role
        for m in re.finditer(
            r"\n\s*([آ-یA-Za-z\u0600-\u06FF][آ-یA-Za-z\u0600-\u06FF\s‌]{2,40})\s*\n\s*(مدیر شعبه|سرپرست فروش شعبه|سرپرست فروش)\s*",
            text,
        ):
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
            role = m.group(2).strip()
            if any(x in pe for x in ("آدرس", "تلفن", "مرکز", "مناطق", "تماس", "درباره")):
                continue
            hits.append(
                {
                    "company": f"بهستان پخش — {city}",
                    "person": f"{pe} ({role} بهستان {city})",
                    "mobile": None,
                    "land": lands[0] if lands else land("02184054000"),
                    "src": "behestan_branch",
                    "url": url,
                }
            )
    return hits


def main():
    hits = []
    hits.extend(scrape_apir())
    hits.extend(scrape_behestan())
    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
