# -*- coding: utf-8 -*-
"""Import Arak (and similar) provincial pharma Excels into person_id_hits.jsonl."""
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


def normalize_land(v: str, default_area: str = "086") -> str | None:
    d = re.sub(r"\D", "", S(v).translate(DIGITS).split("\n")[0])
    if not d or d.startswith("09") or (len(d) == 10 and d.startswith("9")):
        return None
    if d.startswith("98"):
        d = "0" + d[2:]
    if not d.startswith("0"):
        if len(d) == 8:
            d = default_area + d
        elif len(d) == 10:
            d = "0" + d
    return d


def import_arak(path: Path, label: str) -> list[dict]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    headers = [S(c.value).replace("\u200c", "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1))]
    # fuzzy column find
    def col(*names):
        for i, h in enumerate(headers):
            for n in names:
                if n in h:
                    return i
        return None

    i_co = col("نام شرکت", "شرکت")
    i_pe = col("مدیر فروش", "مدیر", "مدیرعامل")
    i_tech = col("مسئول فنی")
    i_tel = col("تلفن")
    i_city = col("شهر")
    hits = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or i_co is None:
            continue
        co = S(r[i_co])
        if not co:
            continue
        city = S(r[i_city]) if i_city is not None else ""
        branch = f"{co} — شعبه {city or 'مرکزی/قم'}"
        land = normalize_land(r[i_tel]) if i_tel is not None else None
        for idx, pe_i in enumerate([i_pe, i_tech]):
            if pe_i is None:
                continue
            pe = S(r[pe_i])
            if not pe:
                continue
            role = "مدیر فروش" if idx == 0 else "مسئول فنی"
            hits.append(
                {
                    "company": branch,
                    "person": f"{pe} — {role} ({label})",
                    "mobile": None,
                    "land": land,
                    "src": label,
                }
            )
    return hits


def import_generic(path: Path, label: str) -> list[dict]:
    """Best-effort import for Excels with company/person/mobile columns."""
    wb = openpyxl.load_workbook(path, data_only=True)
    hits = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        headers = [S(h).lower() for h in rows[0]]
        # map
        def find(*keys):
            for i, h in enumerate(headers):
                for k in keys:
                    if k in h:
                        return i
            return None

        i_co = find("شرکت", "company", "نام")
        i_pe = find("مدیر", "person", "فرد", "مسئول")
        i_mob = find("همراه", "موبایل", "mobile")
        i_tel = find("تلفن", "ثابت", "phone")
        if i_co is None:
            continue
        for r in rows[1:]:
            if not r:
                continue
            co = S(r[i_co]) if i_co < len(r) else ""
            pe = S(r[i_pe]) if i_pe is not None and i_pe < len(r) else ""
            if not co or not pe:
                continue
            mob = to_mobile(r[i_mob]) if i_mob is not None and i_mob < len(r) else None
            land = normalize_land(r[i_tel]) if i_tel is not None and i_tel < len(r) else None
            hits.append(
                {
                    "company": f"{co} — {label}",
                    "person": f"{pe} — مخاطب ({label})",
                    "mobile": mob,
                    "land": land,
                    "src": label,
                }
            )
    return hits


def main() -> None:
    hits: list[dict] = []
    p1 = RAW / "arak_pharma_1401.xlsx"
    p2 = RAW / "arak_pharma_1402.xlsx"
    if p1.exists():
        hits.extend(import_arak(p1, "arak_1401"))
    if p2.exists():
        # try same schema; fallback generic
        try:
            hits.extend(import_arak(p2, "arak_1402"))
        except Exception:
            hits.extend(import_generic(p2, "arak_1402"))
    isf = RAW / "isfahan_pharma.xlsx"
    if isf.exists():
        hits.extend(import_generic(isf, "isfahan_fdo"))

    existing = []
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if line.strip():
                existing.append(json.loads(line))
    seen = {(h.get("company"), h.get("person"), h.get("mobile") or "") for h in existing}
    added = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            key = (h.get("company"), h.get("person"), h.get("mobile") or "")
            if key in seen or not h.get("person"):
                continue
            seen.add(key)
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    print(json.dumps({"batch_hits": len(hits), "appended": added}, ensure_ascii=False))


if __name__ == "__main__":
    main()
