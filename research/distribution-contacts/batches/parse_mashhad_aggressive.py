# -*- coding: utf-8 -*-
"""Aggressively extract all company/person/mobile triples from Mashhad PDF text."""
from __future__ import annotations

import json
import re
from pathlib import Path

ROOT = Path(r"D:\HomeDesign")
TXT = ROOT / "research/distribution-contacts/raw/mashhad_pharma.txt"
OUT = ROOT / "research/distribution-contacts/batches/person_id_hits.jsonl"
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")

HEADERS = {
    "نام شرکت",
    "نام مدير شرکت",
    "نام مدیر شرکت",
    "تلفن ثابت",
    "تلفن همراه",
    "نام مسئول فني",
    "نام مسئول فنی",
    "آدرس شرکت",
    "مجوز زنجيره سرد",
    "مجوز زنجیره سرد",
    "داروی مخدر",
    "جوز دريافت داروهای شيمي درما",
    "مجوز دريافت سرم",
    "مجوز دریافت سرم",
    "مجوز پوشش استان های همجوار",
}


def to_mobile(v: str) -> str | None:
    d = re.sub(r"\D", "", str(v).translate(DIGITS))
    if d.startswith("98") and len(d) >= 12:
        d = "0" + d[2:]
    if len(d) == 10 and d.startswith("9"):
        d = "0" + d
    if len(d) == 11 and d.startswith("09"):
        return d
    return None


def is_header(s: str) -> bool:
    return s.strip() in HEADERS or re.fullmatch(r"[√X×x\s]+", s or "")


def looks_like_company(s: str) -> bool:
    s = s.strip()
    if not s or is_header(s) or len(s) < 3 or len(s) > 80:
        return False
    if not re.search(r"[آ-ی]", s):
        return False
    # company lines usually don't start with آقای/خانم/دکتر alone as whole
    if re.match(r"^(آقای|اقای|خانم|دکتر)\s", s) and to_mobile(s):
        return False
    # address-ish
    if re.match(r"^(آزادی|بلوار|خیابان|خ |کیلومتر|شهرک|توس|منطقه|طرق)", s):
        return False
    if re.fullmatch(r"[\d\-\s/]+", s.translate(DIGITS)):
        return False
    return True


def extract_named_mobiles(blob: str) -> list[tuple[str, str]]:
    blob = blob.translate(DIGITS)
    pairs = []
    for m in re.finditer(
        r"([آابپتثجچحخدذرزژسشصضطظعغفقکگلمنوهیءؤئيك‌\s\.]{2,45}?)(0?9\d{9})",
        blob,
    ):
        name = m.group(1).strip()
        name = re.sub(r"\d{5,}$", "", name).strip(" -–—\t.")
        name = re.sub(r"\s+", " ", name)
        mob = to_mobile(m.group(2))
        if not mob or len(name) < 3:
            continue
        if is_header(name):
            continue
        if sum(c.isdigit() for c in name) > 2:
            continue
        if re.match(r"^(آزادی|بلوار|خیابان|کیلومتر|شهرک|توس)", name):
            continue
        pairs.append((name, mob))
    return pairs


def main() -> None:
    raw = TXT.read_text(encoding="utf-8")
    lines = [ln.strip() for ln in raw.splitlines() if ln.strip()]

    hits = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if not looks_like_company(line):
            i += 1
            continue
        # collect following content until next company-looking line with no mobiles nearby
        blob_parts = []
        j = i + 1
        while j < len(lines) and j <= i + 12:
            nxt = lines[j]
            if is_header(nxt):
                j += 1
                # don't break on headers mid-table; skip
                if nxt in HEADERS and j > i + 2 and not blob_parts:
                    break
                continue
            # next company?
            if (
                looks_like_company(nxt)
                and not re.search(r"0?9\d{9}", nxt.translate(DIGITS))
                and j > i + 1
                and blob_parts
            ):
                break
            blob_parts.append(nxt)
            j += 1
        blob = " ".join(blob_parts)
        pairs = extract_named_mobiles(blob)
        # also mobiles on their own line after a name line
        if not pairs:
            # try company line itself may have glued mobile? rare
            pairs = extract_named_mobiles(line + " " + blob)
        if pairs:
            for idx, (pe, mob) in enumerate(pairs[:3]):  # manager + up to 2 technical
                role = "مدیر شرکت" if idx == 0 else "مسئول فنی/مخاطب"
                hits.append(
                    {
                        "company": f"{line} — شعبه مشهد",
                        "person": f"{pe} — {role} (فهرست پخش سلامت مشهد)",
                        "mobile": mob,
                        "land": None,
                        "src": "mashhad_aggressive",
                    }
                )
        i = max(j, i + 1)

    existing = []
    if OUT.exists():
        for line in OUT.read_text(encoding="utf-8").splitlines():
            if line.strip():
                existing.append(json.loads(line))
    seen = {(h.get("company"), h.get("mobile") or "") for h in existing}
    # also dedupe by mobile alone for mashhad branch to avoid duplicates
    seen_mobs = {h.get("mobile") for h in existing if h.get("mobile") and "مشهد" in str(h.get("company", ""))}
    added = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in hits:
            if not h.get("mobile"):
                continue
            key = (h.get("company"), h.get("mobile"))
            if key in seen or h["mobile"] in seen_mobs:
                continue
            seen.add(key)
            seen_mobs.add(h["mobile"])
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            added += 1
    print(json.dumps({"candidates": len(hits), "appended": added}, ensure_ascii=False))


if __name__ == "__main__":
    main()
