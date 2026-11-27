# -*- coding: utf-8 -*-
"""Import Pegah Bazar Gostar + Setayesh branch managers with unique mobiles."""
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
BLOCK = ("کارشناس فروش", "ویزیتور", "فروشنده")


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


def append_unique_mobile_hits(hits: list[dict]) -> dict:
    existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()] if OUT.exists() else []
    seen_keys = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    seen_mobs = {h.get("mobile") for h in existing if h.get("mobile")}
    added = new_m = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            pe = h.get("person") or ""
            if any(b in pe for b in BLOCK):
                continue
            mob = h.get("mobile")
            if not mob or mob in seen_mobs:
                continue
            key = (h["company"], pe, mob)
            if key in seen_keys:
                continue
            seen_keys.add(key)
            seen_mobs.add(mob)
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
            new_m += 1
    return {"appended": added, "new_unique_mobiles": new_m}


def scrape_pegah() -> list[dict]:
    hits = []
    url = "https://bazar.pegah.ir/%D9%85%D9%86%D8%A7%D8%B7%D9%82/"
    html = fetch(url)
    (RAW / "pegah_bazar_regions.html").write_text(html, encoding="utf-8")
    # table rows
    for m in re.finditer(r"<tr[^>]*>([\s\S]*?)</tr>", html, re.I):
        cells = re.findall(r"<t[dh][^>]*>([\s\S]*?)</t[dh]>", m.group(1), re.I)
        cells = [re.sub(r"<[^>]+>", " ", c) for c in cells]
        cells = [re.sub(r"\s+", " ", c).strip() for c in cells]
        if len(cells) < 8:
            continue
        # expected: منطقه، نام شعبه، آدرس، تلفن، کدپستی، فکس، نام مدیر، پست سازمانی، تلفن همراه
        branch = cells[1] if len(cells) > 1 else ""
        phone = cells[3] if len(cells) > 3 else ""
        manager = cells[6] if len(cells) > 6 else ""
        role = cells[7] if len(cells) > 7 else "مدیر شعبه"
        mobile = cells[8] if len(cells) > 8 else ""
        if not branch or "نام شعبه" in branch:
            continue
        pe = manager.replace("آقای ", "").replace("خانم ", "").strip()
        if not pe or pe in ("نام مدیر شعبه",):
            continue
        if any(b in role for b in BLOCK) or any(b in pe for b in BLOCK):
            continue
        # accept مدیر / سرپرست / مدیریت
        if not any(k in role for k in ("مدیر", "سرپرست", "مدیریت")):
            continue
        mb = to_mobile(mobile)
        if not mb:
            # sometimes mobile without leading 0
            mb = to_mobile("0" + re.sub(r"\D", "", mobile.translate(DIGITS))) if mobile else None
        if not mb:
            continue
        hits.append(
            {
                "company": f"بازارگستر پگاه — {branch}",
                "person": f"{pe} ({role} بازارگستر پگاه)",
                "mobile": mb,
                "land": land(phone),
                "src": "pegah_bazar",
            }
        )
    # fallback text parse for mobiles near names
    text = re.sub(r"<[^>]+>", "\n", html)
    text = re.sub(r"\n+", "\n", text).translate(DIGITS)
    (RAW / "pegah_bazar_regions_text.txt").write_text(text, encoding="utf-8")
    print({"pegah_table_hits": len(hits)})
    return hits


def scrape_setayesh() -> list[dict]:
    hits = []
    try:
        html = fetch("https://setayesh-pakhsh.com/%d8%b4%d8%b9%d8%a8/")
        (RAW / "setayesh_branches.html").write_text(html, encoding="utf-8")
        text = re.sub(r"<[^>]+>", "\n", html)
        text = re.sub(r"\n+", "\n", text).translate(DIGITS)
        for m in re.finditer(
            r"شعبه\s*([^\n]{2,40})\s*\n([\s\S]{0,300}?)(0?9\d{9})",
            text,
        ):
            city = re.sub(r"\s+", " ", m.group(1)).strip()
            mb = to_mobile(m.group(3))
            if not mb:
                continue
            hits.append(
                {
                    "company": f"ستایش پخش — شعبه {city}",
                    "person": f"مدیر/مسئول شعبه ستایش {city}",
                    "mobile": mb,
                    "land": None,
                    "src": "setayesh",
                }
            )
    except Exception as e:
        print({"setayesh_error": str(e)})
    print({"setayesh_hits": len(hits)})
    return hits


EXTRA = [
    ("ایران پخش / ایرانیان پخش نگین بارثاوا — مشهد", "محمد صباغیان", "مدیرعامل", "09045452150", "05137133500"),
]


def main():
    hits = []
    hits.extend(scrape_pegah())
    hits.extend(scrape_setayesh())
    for co, pe, role, mob, ld in EXTRA:
        hits.append(
            {
                "company": co,
                "person": f"{pe} ({role})",
                "mobile": to_mobile(mob),
                "land": land(ld),
                "src": "extra_mgr_mobile",
            }
        )
    stats = append_unique_mobile_hits(hits)
    print(json.dumps({"prepared": len(hits), **stats}, ensure_ascii=False))


if __name__ == "__main__":
    main()
