# -*- coding: utf-8 -*-
"""Import Arak 1402 (CEO+mobile) and Isfahan FDO (manager+phone) Excels."""
from __future__ import annotations

import json
import re
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\HomeDesign")
RAW = ROOT / "research/distribution-contacts/raw"
OUT = ROOT / "research/distribution-contacts/batches/person_id_hits.jsonl"
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


def S(v) -> str:
    return "" if v is None else str(v).strip()


def to_mobile(v: str) -> str | None:
    d = re.sub(r"\D", "", S(v).translate(DIGITS))
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    if len(d) == 11 and d.startswith("09"):
        return d
    return None


def all_mobiles(v: str) -> list[str]:
    out = []
    for part in re.split(r"[\n,/|]+", S(v)):
        m = to_mobile(part)
        if m and m not in out:
            out.append(m)
    # also scan whole string
    blob = re.sub(r"\D", "", S(v).translate(DIGITS))
    for m in re.finditer(r"0?9\d{9}", blob):
        mm = to_mobile(m.group(0))
        if mm and mm not in out:
            out.append(mm)
    return out


def normalize_land(v: str, default_area: str = "086") -> str | None:
    first = S(v).split("\n")[0]
    first = re.sub(r"(?i)تلفن\s*:?", "", first)
    first = re.sub(r"(?i)نمابر.*", "", first)
    d = re.sub(r"\D", "", first.translate(DIGITS))
    if not d or d.startswith("09") or (len(d) == 10 and d.startswith("9")):
        return None
    if d.startswith("98"):
        d = "0" + d[2:]
    if not d.startswith("0"):
        if len(d) in (7, 8):
            d = default_area + d
        elif len(d) == 10:
            d = "0" + d
    return d


def split_people(v: str) -> list[str]:
    parts = re.split(r"[\n/;|]+", S(v))
    out = []
    for p in parts:
        p = re.sub(r"مدیر فروش.*", "", p).strip()
        p = re.sub(r"\s+", " ", p).strip(" -–—")
        if p and len(p) >= 3:
            out.append(p)
    return out


def import_arak1402() -> list[dict]:
    path = RAW / "arak_pharma_1402.xlsx"
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    hits = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        co = S(r[1])
        if not co:
            continue
        city = S(r[2]) or "مرکزی"
        land = normalize_land(r[3], "086" if city == "اراک" else "025")
        people = split_people(r[5])
        mobiles = all_mobiles(r[6])
        tech = split_people(r[7]) if len(r) > 7 else []
        company = f"{co} — شعبه {city}"
        if people:
            pe = people[0]
            mob = mobiles[0] if mobiles else None
            hits.append(
                {
                    "company": company,
                    "person": f"{pe} — مدیرعامل/مدیر فروش (arak_1402)",
                    "mobile": mob,
                    "land": land,
                    "src": "arak_1402_ceo",
                }
            )
            # extra people+mobiles as additional contacts
            for i, pe2 in enumerate(people[1:]):
                mob2 = mobiles[i + 1] if len(mobiles) > i + 1 else None
                hits.append(
                    {
                        "company": company,
                        "person": f"{pe2} — مدیر فروش/مخاطب (arak_1402)",
                        "mobile": mob2,
                        "land": land,
                        "src": "arak_1402_extra",
                    }
                )
            for i, m in enumerate(mobiles[1:] if len(people) == 1 else []):
                hits.append(
                    {
                        "company": company,
                        "person": f"{pe} — مخاطب اضافی (arak_1402)",
                        "mobile": m,
                        "land": land,
                        "src": "arak_1402_extra_mob",
                    }
                )
        for t in tech:
            hits.append(
                {
                    "company": company,
                    "person": f"{t} — مسئول فنی (arak_1402)",
                    "mobile": None,
                    "land": land,
                    "src": "arak_1402_tech",
                }
            )
    return hits


def import_isfahan() -> list[dict]:
    path = RAW / "isfahan_pharma.xlsx"
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [S(c.value) for c in next(ws.iter_rows(min_row=1, max_row=1))]
    # indices from known headers
    # نام شرکت, نام شعبه, نام شرکت/مدیر, آدرس, نام مسئول فنی, شماره تماس مسئول فنی, نام مدیر پخش, شماره تماس مدیر
    hits = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r:
            continue
        co = S(r[2]) or S(r[3])
        branch = S(r[3]) or "اصفهان"
        mgr_combo = S(r[4])  # "شرکت - شخص"
        tech = S(r[12]) if len(r) > 12 else ""
        tech_tel = S(r[13]) if len(r) > 13 else ""
        mgr = S(r[17]) if len(r) > 17 else ""
        mgr_tel = S(r[18]) if len(r) > 18 else ""
        addr = S(r[9]) if len(r) > 9 else ""
        land = None
        for m in re.finditer(r"0?\d{7,11}", addr.translate(DIGITS)):
            land = normalize_land(m.group(0), "031")
            if land:
                break

        person = mgr
        if not person and mgr_combo and " - " in mgr_combo:
            person = mgr_combo.split(" - ", 1)[-1].strip()
        elif not person and mgr_combo and " – " in mgr_combo:
            person = mgr_combo.split(" – ", 1)[-1].strip()
        elif not person and mgr_combo:
            # try last token after dash variants
            parts = re.split(r"\s*[-–—]\s*", mgr_combo)
            if len(parts) >= 2:
                person = parts[-1].strip()
        if not co or not person:
            continue
        company = f"{co} — شعبه اصفهان"
        mobs = all_mobiles(mgr_tel)
        hits.append(
            {
                "company": company,
                "person": f"{person} — مدیر پخش شعبه اصفهان (fdo.mui)",
                "mobile": mobs[0] if mobs else None,
                "land": land,
                "src": "isfahan_fdo_mgr",
            }
        )
        if tech:
            tm = all_mobiles(tech_tel)
            hits.append(
                {
                    "company": company,
                    "person": f"{tech} — مسئول فنی شعبه اصفهان (fdo.mui)",
                    "mobile": tm[0] if tm else None,
                    "land": land,
                    "src": "isfahan_fdo_tech",
                }
            )
    return hits


def main() -> None:
    hits = import_arak1402() + import_isfahan()
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
            if key in seen or not h.get("person"):
                continue
            seen.add(key)
            if h.get("mobile"):
                with_mob += 1
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    print(json.dumps({"hits": len(hits), "appended": added, "with_mobile": with_mob}, ensure_ascii=False))


if __name__ == "__main__":
    main()
