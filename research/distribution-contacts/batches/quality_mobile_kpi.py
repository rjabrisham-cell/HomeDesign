# -*- coding: utf-8 -*-
"""Quality KPI: unique mobiles linked to company and/or mid/senior manager."""
from __future__ import annotations

import json
import re
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
CANDIDATES = [
    ROOT / "Distribution-Contact-Directory-FA.xlsx",
    Path("artifacts/Distribution-Contact-Directory-FA-updated.xlsx"),
    Path("artifacts/Distribution-Contact-Directory-FA.xlsx"),
]
XLSX = next((p for p in CANDIDATES if p.exists()), CANDIDATES[0])
OUT = ROOT / "batches" / "quality_mobile_kpi.json"

DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

SALES_ONLY = ("کارشناس فروش", "ویزیتور", "فروشنده", "کارشناس ویزیت")
MANAGER_OK = (
    "مدیرعامل",
    "مدیر عامل",
    "مدیر فروش",
    "مدیر بازرگانی",
    "معاون",
    "رئیس",
    "مدیر شعبه",
    "مدیر مرکز",
    "مدیر کل",
    "هیئت مدیره",
    "هیات مدیره",
    "خزانه",
    "دبیر",
    "سرپرست شعبه",
    "سرپرست فروش شعبه",
    "مدیر منطقه",
    "مدیر ارشد",
    "عضو هیئت",
    "بازرس",
)


def to_mobile(v):
    if not v:
        return None
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    return d if len(d) == 11 and d.startswith("09") else None


def is_sales_only(pe: str) -> bool:
    pe = pe or ""
    if any(k in pe for k in MANAGER_OK):
        return False
    return any(k in pe for k in SALES_ONLY)


def is_managerish(pe: str) -> bool:
    pe = pe or ""
    if is_sales_only(pe):
        return False
    return any(k in pe for k in MANAGER_OK)


def main():
    wb = load_workbook(XLSX)
    ws = wb.active
    unique = {}  # mobile -> best link
    sales_linked = set()
    for r in ws.iter_rows(min_row=2, values_only=True):
        co, pe, mob, land = (r + (None, None, None, None))[:4]
        co = str(co).strip() if co else ""
        pe = str(pe).strip() if pe else ""
        mcol = str(mob).strip() if mob else ""
        for raw in mcol.splitlines() if mcol else []:
            m = to_mobile(raw)
            if not m:
                continue
            if is_sales_only(pe):
                sales_linked.add(m)
            entry = unique.get(m)
            score = 0
            if co:
                score += 2
            if is_managerish(pe):
                score += 3
            elif pe and not is_sales_only(pe):
                score += 1
            cand = {"mobile": m, "company": co, "person": pe, "score": score, "sales_only_person": is_sales_only(pe)}
            if not entry or cand["score"] > entry["score"]:
                unique[m] = cand

    # Quality: unique mobile with company OR non-sales person (prefer manager)
    quality = []
    for m, e in unique.items():
        if e["company"] or (e["person"] and not e["sales_only_person"]):
            quality.append(e)

    mgr = sum(1 for e in quality if is_managerish(e["person"]))
    co_only = sum(1 for e in quality if e["company"] and not e["person"])
    other_pe = sum(1 for e in quality if e["person"] and not is_managerish(e["person"]) and not e["sales_only_person"])

    stats = {
        "unique_mobiles_all": len(unique),
        "unique_mobiles_quality": len(quality),
        "target": 5000,
        "gap": max(0, 5000 - len(quality)),
        "with_manager_person": mgr,
        "company_mobile_no_person": co_only,
        "other_non_sales_person": other_pe,
        "mobiles_also_seen_with_sales_label": len(sales_linked),
        "note": "Quality = unique 09 mobile linked to company and/or non-sales person; prefer mid/senior manager. Existing sales rows kept, not deleted.",
    }
    OUT.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(stats, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
