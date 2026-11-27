# -*- coding: utf-8 -*-
"""Scrape Adorateb location pages + refresh Karen branch show pages."""
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


def fetch(url: str, timeout: int = 25) -> str:
    req = urllib.request.Request(url, headers=UA)
    return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")


def html_to_text(html: str) -> str:
    text = re.sub(r"<script[\s\S]*?</script>", " ", html, flags=re.I)
    text = re.sub(r"<style[\s\S]*?</style>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "\n", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n+", "\n", text).translate(DIGITS)


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


def scrape_adorateb() -> list[dict]:
    hits = []
    try:
        index = fetch("https://adorateb.com/%D9%85%D8%B1%D8%A7%DA%A9%D8%B2-%D9%BE%D8%AE%D8%B4/")
        (RAW / "adorateb_centers.html").write_text(index, encoding="utf-8")
    except Exception as e:
        print({"adorateb_index_error": str(e)})
        return hits

    links = sorted(
        set(
            re.findall(
                r'href="(https://adorateb\.com/location/[^"]+|/%D9%85%D8%B1%DA%A9%D8%B2-[^"]+|/location/[^"]+)"',
                index,
            )
        )
    )
    # also relative persian paths
    more = re.findall(r'href="(/[^\"]*%D9%85%D8%B1%DA%A9%D8%B2[^\"]*)"', index)
    for m in more:
        if m not in links:
            links.append(m)
    abs_links = []
    for l in links:
        if l.startswith("http"):
            abs_links.append(l)
        else:
            abs_links.append("https://adorateb.com" + l)
    abs_links = sorted(set(abs_links))
    print({"adorateb_links": len(abs_links)})

    for url in abs_links[:40]:
        try:
            html = fetch(url)
        except Exception:
            continue
        text = html_to_text(html)
        title_m = re.search(r"مرکز پخش دارو\s+([^\n]+)", text)
        city = title_m.group(1).strip() if title_m else "نامشخص"
        # person often first bullet-like name near phones
        pe = None
        # look for Persian full name lines before phone
        for m in re.finditer(r"\n\s*([آ-یA-Za-z\u0600-\u06FF]{2,}(?:\s+[آ-یA-Za-z\u0600-\u06FF]{2,}){1,3})\s*\n", text):
            cand = m.group(1).strip()
            if any(x in cand for x in ("مرکز", "شرکت", "آدورا", "تلفن", "آدرس", "استان", "ساعت", "دارو")):
                continue
            if len(cand) < 6 or len(cand) > 40:
                continue
            pe = cand
            break
        lands = []
        for mm in re.finditer(r"(?<!\d)(0\d{2,3}\d{7,8})(?!\d)", text):
            d = mm.group(1)
            if not d.startswith("09") and d not in lands and not d.startswith("021883"):
                lands.append(d)
        mobs = []
        for mm in re.finditer(r"(?<!\d)(0?9\d{9})(?!\d)", text):
            mb = to_mobile(mm.group(1))
            if mb and mb not in mobs:
                mobs.append(mb)
        if pe or mobs:
            hits.append(
                {
                    "company": f"آدوراطب — مرکز {city}",
                    "person": f"{pe or 'مخاطب مرکز'} (آدوراطب {city})",
                    "mobile": mobs[0] if mobs else None,
                    "land": lands[0] if lands else None,
                    "src": "adorateb_scrape",
                    "url": url,
                }
            )
    return hits


def scrape_karen() -> list[dict]:
    hits = []
    # known branch ids from prior JSON + a wider sweep
    ids = list(range(1, 40))
    for bid in ids:
        url = f"https://karendistribution.com/branch/show/{bid}"
        try:
            html = fetch(url, timeout=15)
        except Exception:
            continue
        if "404" in html and "شعبه" not in html and "موبایل" not in html:
            continue
        text = html_to_text(html)
        title_m = re.search(r"\n\s*([آ-یA-Za-z\u0600-\u06FF]{2,30})\s*\n", text)
        city = title_m.group(1).strip() if title_m else f"شعبه {bid}"
        pe = None
        m = re.search(r"(?:نام و نام خانوادگی|مدیر شعبه)\s*\n?\s*([آ-یA-Za-z\u0600-\u06FF\s‌]{3,40})", text)
        if m:
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
            pe = pe.replace("آقای ", "").replace("خانم ", "")
        mb = None
        mm = re.search(r"موبایل\s*\n?\s*(0?9\d{9})", text)
        if mm:
            mb = to_mobile(mm.group(1))
        ld = None
        lm = re.search(r"شماره تماس\s*\n?\s*(0\d{2,3}\d{7,8})", text)
        if lm:
            ld = land(lm.group(1))
        if pe or mb:
            hits.append(
                {
                    "company": f"پخش کارن — شعبه {city}",
                    "person": f"{pe or 'مدیر شعبه'} (مدیر شعبه کارن {city})",
                    "mobile": mb,
                    "land": ld,
                    "src": "karen_refresh",
                    "url": url,
                }
            )
    (RAW / "karen_refresh_hits.json").write_text(json.dumps(hits, ensure_ascii=False, indent=2), encoding="utf-8")
    return hits


def main():
    hits = []
    hits.extend(scrape_adorateb())
    hits.extend(scrape_karen())
    added, wm = append_hits(hits)
    print(
        json.dumps(
            {"scraped": len(hits), "appended": added, "with_mobile": wm},
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
