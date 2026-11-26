# -*- coding: utf-8 -*-
"""Scrape APIR provincial branch boards nationwide."""
from __future__ import annotations

import json
import re
import urllib.request
from pathlib import Path

OUT = Path(r"D:\HomeDesign\research\distribution-contacts\batches\person_id_hits.jsonl")
RAW = Path(r"D:\HomeDesign\research\distribution-contacts\raw")
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def land(v: str | None) -> str | None:
    if not v:
        return None
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if not d or d.startswith("09") or (len(d) == 10 and d.startswith("9")):
        return None
    if not d.startswith("0") and len(d) == 10:
        d = "0" + d
    if len(d) == 8:
        d = "021" + d
    return d if d.startswith("0") and 10 <= len(d) <= 11 else None


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


def scrape() -> list[dict]:
    hits = []
    for i in range(1, 35):
        url = f"https://apir.ir/branches/branch-{i}"
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=20) as r:
                html = r.read().decode("utf-8", "ignore")
        except Exception:
            continue
        # province title
        title_m = re.search(r"<h1[^>]*>\s*([^<]+)\s*</h1>", html, re.I)
        province = re.sub(r"\s+", " ", title_m.group(1)).strip() if title_m else f"شعبه {i}"
        # phones
        lands = []
        mobs = []
        for mm in re.finditer(r"(?<!\d)(0\d{2}\d{7,8}|0?9\d{9})(?!\d)", html.translate(DIGITS)):
            d = mm.group(1)
            if d.startswith("09") or (len(d) == 10 and d.startswith("9")):
                mb = to_mobile(d)
                if mb and mb not in mobs:
                    mobs.append(mb)
            else:
                ld = land(d)
                if ld and ld not in lands:
                    lands.append(ld)
        # board table rows: name | role
        # look for table cells pairs
        rows = re.findall(r"<tr[^>]*>([\s\S]*?)</tr>", html, re.I)
        people = []
        for row in rows:
            cells = [
                re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", c)).strip()
                for c in re.findall(r"<t[dh][^>]*>([\s\S]*?)</t[dh]>", row, re.I)
            ]
            if len(cells) >= 2:
                name, role = cells[0], cells[1]
                if name in ("نام و نام خانوادگی", "نام") or not name:
                    continue
                if "هیئت" in name and "مدیر" in name:
                    continue
                people.append((name, role))
        # also catch card-style names if table empty
        if not people:
            for m in re.finditer(
                r"(آقای|خانم|سرکار خانوم)\s+([^<\n]{3,50})\s*</[^>]+>\s*<[^>]+>\s*([^<\n]{3,40})",
                html,
            ):
                people.append((m.group(1) + " " + m.group(2).strip(), m.group(3).strip()))
        for name, role in people:
            pe = f"{name} — {role} (انجمن پخش {province} / apir)"
            hits.append(
                {
                    "company": f"انجمن ملی صنعت پخش — {province}",
                    "person": pe,
                    "mobile": mobs[0] if mobs else None,
                    "land": lands[0] if lands else None,
                    "src": "apir_branch_scrape",
                    "branch_id": i,
                }
            )
            # only attach mobile to first person of branch to avoid fake triples
            mobs = []
    (RAW / "apir_branches.json").write_text(
        json.dumps(hits, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return hits


def main() -> None:
    hits = scrape()
    existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()]
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = wm = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            key = (h.get("company"), h.get("person"), h.get("mobile") or "")
            if key in seen:
                continue
            seen.add(key)
            if h.get("mobile"):
                wm += 1
            out = {
                "company": h["company"],
                "person": h["person"],
                "mobile": h.get("mobile"),
                "land": h.get("land"),
                "src": h.get("src"),
            }
            f.write(json.dumps(out, ensure_ascii=False) + "\n")
            added += 1
    print(json.dumps({"batch": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
