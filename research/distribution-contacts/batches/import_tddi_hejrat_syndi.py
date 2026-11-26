# -*- coding: utf-8 -*-
"""Import TDDI provincial agents, Hejrat Shiraz+, Syndipharma managers."""
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
    if not d.startswith("0") and len(d) == 10:
        d = "0" + d
    return d if d.startswith("0") and 10 <= len(d) <= 11 else None


# From tddico.com provincial agents table (animal drug distributors — still B2B distribution)
TDDI = [
    ("کیمیا وارنا — کرج/البرز", "علیپور — مسئول نمایندگی (tddico)", None, "02634704164"),
    ("هوراندیش خزر — گیلان", "بابازاده — مسئول نمایندگی (tddico)", "09158763651", "01333581449"),
    ("پیام‌آوران سلامت هیرکان — گلستان", "میرکریمی — مسئول نمایندگی (tddico)", None, "01732127499"),
    ("وتاک دارو شمال — مازندران", "لطفی‌نژاد — مسئول نمایندگی (tddico)", "09113266587", None),
    ("آمارد داروآوران شمال — مازندران", "پورشعبان — مسئول نمایندگی (tddico)", "094616173471", "01144231305"),
]

HEJRAT_KNOWN = [
    ("شیراز", "محمدرضا حسن‌زاده", "07138247775"),
]

# Syndipharma affiliated distributors (manager + land) — curated high-value HQ
SYNDI = [
    ("پخش البرز", "عقیل آرین‌نژاد — مدیرعامل (syndipharma)", "02166707727"),
]


def scrape_hejrat() -> list[dict]:
    hits = []
    for i in range(1, 35):
        url = f"https://hejratco.com/DistributeCenter/Details/{i}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=20) as r:
                html = r.read().decode("utf-8", "ignore")
        except Exception:
            continue
        # city from "اطلاعات مدیر مرکز X"
        cm = re.search(r"اطلاعات مدیر مرکز\s*([^\n<]{2,40})", html)
        city = cm.group(1).strip() if cm else f"مرکز {i}"
        pe = None
        m = re.search(
            r"اطلاعات مدیر مرکز[\s\S]{0,800}?<(?:h[1-6]|strong|b)[^>]*>\s*([^<]{3,60})\s*</",
            html,
            re.I,
        )
        if m:
            pe = re.sub(r"\s+", " ", m.group(1)).strip()
        if not pe or pe in ("نام مدیر", "-", "تاریخ تولد"):
            # try first h5 after section
            parts = re.findall(r"<h5[^>]*>\s*([^<]{3,60})\s*</h5>", html, re.I)
            for p in parts:
                p = re.sub(r"\s+", " ", p).strip()
                if p and not p.startswith("13") and "نام" not in p:
                    pe = p
                    break
        lands = []
        for mm in re.finditer(r"(?<!\d)(0\d{2}\d{7,8})(?!\d)", html.translate(DIGITS)):
            d = mm.group(1)
            if not d.startswith("09") and d not in lands:
                lands.append(d)
        if pe:
            hits.append(
                {
                    "company": f"پخش هجرت — مرکز {city}",
                    "person": f"{pe} — مدیر مرکز (hejratco)",
                    "mobile": None,
                    "land": lands[0] if lands else None,
                    "src": "hejrat_scrape2",
                }
            )
    (RAW / "hejrat_centers2.json").write_text(
        json.dumps(hits, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return hits


def scrape_tddi_page() -> list[dict]:
    url = "https://tddico.com/%d9%86%d9%85%d8%a7%db%8c%d9%86%d8%af%da%af%db%8c-%d9%be%d8%ae%d8%b4-%d8%a7%d8%b3%d8%aa%d8%a7%d9%86%db%8c/"
    hits = []
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=25) as r:
            html = r.read().decode("utf-8", "ignore")
    except Exception as e:
        print("tddi fetch fail", e)
        return hits
    # parse table rows roughly
    rows = re.findall(r"<tr[^>]*>([\s\S]*?)</tr>", html, re.I)
    for row in rows:
        cells = [re.sub(r"<[^>]+>", " ", c) for c in re.findall(r"<t[dh][^>]*>([\s\S]*?)</t[dh]>", row, re.I)]
        cells = [re.sub(r"\s+", " ", c).strip() for c in cells]
        if len(cells) < 5:
            continue
        # expect: idx, province, company, person, address, phone
        province = cells[1] if len(cells) > 1 else ""
        company = cells[2] if len(cells) > 2 else ""
        person = cells[3] if len(cells) > 3 else ""
        phone_blob = " ".join(cells[4:])
        if not company or not person or person in ("مسئول نمایندگی", "نام شرکت"):
            continue
        mobs = []
        for mm in re.finditer(r"(?<!\d)(0?9\d{9}|9\d{9})(?!\d)", phone_blob.translate(DIGITS)):
            mb = to_mobile(mm.group(1))
            if mb and mb not in mobs:
                mobs.append(mb)
        lands = []
        for mm in re.finditer(r"(?<!\d)(0\d{2}\d{7,8})(?!\d)", phone_blob.translate(DIGITS)):
            d = mm.group(1)
            if not d.startswith("09") and d not in lands:
                lands.append(d)
        hits.append(
            {
                "company": f"{company} — {province} (نمایندگی پخش tddico)",
                "person": f"{person} — مسئول نمایندگی استانی",
                "mobile": mobs[0] if mobs else None,
                "land": lands[0] if lands else None,
                "src": "tddi_provincial",
            }
        )
    (RAW / "tddi_agents.json").write_text(
        json.dumps(hits, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return hits


def main() -> None:
    hits: list[dict] = []
    for co, pe, mob, tel in TDDI:
        # fix weird mobile with extra digit
        mb = to_mobile(mob) if mob else None
        if mob and not mb and mob.startswith("09") and len(re.sub(r"\D", "", mob)) == 12:
            # take first 11
            d = re.sub(r"\D", "", mob)[:11]
            mb = to_mobile(d)
        hits.append(
            {
                "company": co,
                "person": pe,
                "mobile": mb,
                "land": land(tel),
                "src": "tddi_curated",
            }
        )
    for city, pe, tel in HEJRAT_KNOWN:
        hits.append(
            {
                "company": f"پخش هجرت — مرکز {city}",
                "person": f"{pe} — مدیر مرکز (hejratco)",
                "mobile": None,
                "land": land(tel),
                "src": "hejrat_known",
            }
        )
    for co, pe, tel in SYNDI:
        hits.append(
            {
                "company": co,
                "person": pe,
                "mobile": None,
                "land": land(tel),
                "src": "syndipharma",
            }
        )
    hits.extend(scrape_hejrat())
    hits.extend(scrape_tddi_page())

    existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()]
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = 0
    with_mob = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            if not h.get("person"):
                continue
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
