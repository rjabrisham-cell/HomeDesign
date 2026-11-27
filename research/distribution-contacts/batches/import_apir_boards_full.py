# -*- coding: utf-8 -*-
"""Full scrape of APIR provincial board tables + remaining logiscm CEOs."""
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


ROLE_WORDS = (
    "رئیس",
    "نایب",
    "خزانه",
    "عضو",
    "بازرس",
    "دبیر",
    "سرکار",
)


def parse_board_html(html: str, url: str) -> list[dict]:
    hits = []
    # province title
    m = re.search(r"استان\s+([^\n<]{2,40})", html)
    if not m:
        m = re.search(r"انجمن\s+([^\n<]{2,40})", html)
    province = re.sub(r"<[^>]+>", "", m.group(1)).strip() if m else "استانی"
    province = re.sub(r"\s+", " ", province)

    lands = []
    text_digits = html.translate(DIGITS)
    for mm in re.finditer(r"(?<!\d)(0\d{2,3}\d{7,8})(?!\d)", text_digits):
        d = mm.group(1)
        if not d.startswith("09") and d not in lands:
            lands.append(d)
    # also pattern like 44009809-021
    for mm in re.finditer(r"(?<!\d)(\d{6,8})\s*[-–]\s*(0?\d{2,3})(?!\d)", text_digits):
        local, area = mm.group(1), mm.group(2)
        if not area.startswith("0"):
            area = "0" + area
        d = area + local
        ld = land(d)
        if ld and ld not in lands:
            lands.append(ld)

    # table rows: <td>آقای NAME</td><td>ROLE</td>
    for m in re.finditer(
        r"<tr[^>]*>\s*<td[^>]*>\s*(?:آقای|خانم|سرکار خانوم|سرکار خانم)?\s*([^<]{3,60})\s*</td>\s*<td[^>]*>\s*([^<]{2,60})\s*</td>",
        html,
        re.I | re.S,
    ):
        pe = re.sub(r"\s+", " ", m.group(1)).strip()
        role = re.sub(r"\s+", " ", m.group(2)).strip()
        if pe in ("نام و نام خانوادگی", "سمت") or role in ("نام و نام خانوادگی", "سمت"):
            continue
        if not any(w in role for w in ROLE_WORDS) and "هیئت" not in role and "هیات" not in role:
            # still accept if looks like board role
            if len(role) < 3:
                continue
        pe = pe.replace("آقای ", "").replace("خانم ", "").replace("سرکار خانوم ", "").replace("سرکار خانم ", "")
        if len(pe) < 3:
            continue
        hits.append(
            {
                "company": f"انجمن ملی صنعت پخش — {province}",
                "person": f"{pe} ({role} انجمن {province})",
                "mobile": None,
                "land": lands[0] if lands else None,
                "src": "apir_board_full",
                "url": url,
            }
        )
    return hits


def scrape_apir() -> list[dict]:
    hits = []
    # discover links from index
    try:
        index = fetch("https://www.apir.ir/branches")
        (RAW / "apir_branches_index.html").write_text(index, encoding="utf-8")
        links = set(re.findall(r'href="([^"]*branches/branch-\d+[^"]*)"', index))
    except Exception as e:
        print({"apir_index_error": str(e)})
        links = set()
    urls = []
    for l in links:
        if l.startswith("http"):
            urls.append(l)
        else:
            urls.append("https://www.apir.ir" + l if l.startswith("/") else "https://apir.ir/" + l)
    # also sweep 1..40
    for n in range(1, 41):
        urls.append(f"https://apir.ir/branches/branch-{n}")
        urls.append(f"https://www.apir.ir/branches/branch-{n}")
    urls = sorted(set(urls))
    print({"apir_urls": len(urls)})
    for url in urls:
        try:
            html = fetch(url, timeout=15)
        except Exception:
            continue
        if "هیئت مدیره" not in html and "نام و نام خانوادگی" not in html:
            continue
        page_hits = parse_board_html(html, url)
        if page_hits:
            print({"url": url, "hits": len(page_hits)})
            hits.extend(page_hits)
            slug = re.sub(r"[^\w]+", "_", url)[-60:]
            (RAW / f"apir_{slug}.html").write_text(html, encoding="utf-8")
    return hits


def remaining_logiscm() -> list[dict]:
    path = ROOT / "queues" / "logiscm_rows.json"
    if not path.exists():
        return []
    rows = json.loads(path.read_text(encoding="utf-8"))
    hits = []
    for r in rows:
        co = (r.get("company") or "").strip()
        pe = (r.get("ceo") or "").strip()
        phone = r.get("phone") or ""
        if not co or not pe:
            continue
        mb = to_mobile(phone)
        ld = None if mb else land(phone)
        # international +98 land often
        if not ld and not mb:
            d = re.sub(r"\D", "", str(phone).translate(DIGITS))
            if d.startswith("98"):
                d = "0" + d[2:]
            ld = land(d)
            mb = to_mobile(d)
        hits.append(
            {
                "company": co,
                "person": f"{pe} (مدیرعامل logiscm)",
                "mobile": mb,
                "land": ld,
                "src": "logiscm_remaining",
            }
        )
    return hits


EXTRA = [
    {
        "company": "آمادپخش",
        "person": "مرسده امیدی (مدیرعامل آمادپخش)",
        "mobile": to_mobile("09044880809"),
        "land": land("02188726405"),
        "src": "amadpakhsh",
    },
    {
        "company": "آمادپخش",
        "person": "مجتبی سلیمانی سدهی (رئیس هیئت مدیره آمادپخش)",
        "mobile": to_mobile("09033306123"),
        "land": None,
        "src": "amadpakhsh",
    },
    {
        "company": "آمادپخش",
        "person": "مصطفی سلیمانی سدهی (عضو هیئت مدیره آمادپخش)",
        "mobile": None,
        "land": land("02188726405"),
        "src": "amadpakhsh",
    },
]


def main():
    hits = []
    hits.extend(scrape_apir())
    hits.extend(remaining_logiscm())
    hits.extend(EXTRA)
    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
