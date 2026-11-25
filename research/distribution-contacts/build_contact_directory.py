# -*- coding: utf-8 -*-
"""Build Distribution-Contact-Directory-FA.xlsx from uploaded lead workbooks + APIR."""
from __future__ import annotations

import json
import re
import subprocess
from collections import defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

UPLOADS = Path(r"C:\Users\rjabr\.cursor\projects\D-HomeDesign\uploads")
OUT_DIR = Path(r"D:\HomeDesign\artifacts")
OUT_DIR.mkdir(parents=True, exist_ok=True)
LOCAL = Path(r"D:\HomeDesign\research\distribution-contacts")
LOCAL.mkdir(parents=True, exist_ok=True)

FILES = {
    "workbook_full": UPLOADS / "workbook_full_df20.xlsx",
    "info_companies": UPLOADS / "_______________1438.xlsx",  # اطلاعات شرکتها
    "brand_contact_a": UPLOADS / "_____________________________9bb0.xlsx",  # اطلاعات تماس برندها
    "brand_b": UPLOADS / "________________bbdb.xlsx",  # برندها و شرکتها
    "customer1399": UPLOADS / "Customer1399_d14f.xlsx",
}

YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(name="Vazirmatn", bold=True, color="FFFFFF", size=12)
BODY_FONT = Font(name="Vazirmatn", size=11)
THIN = Border(
    left=Side(style="thin", color="D0D0D0"),
    right=Side(style="thin", color="D0D0D0"),
    top=Side(style="thin", color="D0D0D0"),
    bottom=Side(style="thin", color="D0D0D0"),
)

# Arabic/Persian digits
DIGITS = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
# remove tatweel, diacritics-ish, unify ye/kaf
_NORM_RE = re.compile(r"[\u0640\u064b-\u065f\u0670\u06d6-\u06ed]")


def norm_name(s) -> str:
    if s is None:
        return ""
    t = str(s).strip().translate(DIGITS)
    t = _NORM_RE.sub("", t)
    t = t.replace("ي", "ی").replace("ك", "ک").replace("ۀ", "ه").replace("ة", "ه")
    t = re.sub(r"\s+", " ", t)
    t = t.replace("آقاي", "آقای").replace("اقای", "آقای").replace("خانم ", "خانم ")
    # strip common prefixes for company match
    return t


def company_key(s: str) -> str:
    t = norm_name(s).lower()
    for p in ("شرکت ", "شركت ", "شرکتهاي ", "موسسه ", "فروشگاه ", "پخش ", "گروه "):
        if t.startswith(p.strip()) or t.startswith(p):
            t = t[len(p) :].strip() if t.startswith(p) else t
    t = re.sub(r"[\(\)\[\]\{\}«»\"'،,.\-_/\\|]+", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t


def person_key(s: str) -> str:
    t = norm_name(s)
    for p in ("آقای ", "اقای ", "خانم ", "مهندس ", "دکتر ", "سرکار خانم "):
        if t.startswith(p):
            t = t[len(p) :].strip()
    return t


def format_phone(raw) -> str | None:
    if raw is None:
        return None
    s = str(raw).strip().translate(DIGITS)
    if not s or s.lower() in ("none", "nan", "-", "—"):
        return None
    # keep as historical text if contains range like 021-4400 or 44009809-2
    if re.search(r"[a-zA-Z@]", s):
        return None  # email etc.
    # Excel may store mobile as int without leading 0
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        n = int(raw)
        s = str(n)
        if len(s) == 10 and s.startswith("9"):
            s = "0" + s
        elif len(s) == 8:
            # landline without area — keep as-is (incomplete lead)
            pass
        elif len(s) == 10 and s.startswith("21"):
            s = "0" + s
    s = s.replace(" ", "").replace("\u200c", "")
    # normalize separators for storage but preserve incomplete ranges
    s2 = s.replace("–", "-").replace("—", "-")
    # if pure digits
    digits = re.sub(r"\D", "", s2)
    if not digits:
        return None
    if len(digits) == 10 and digits.startswith("9"):
        return "0" + digits
    if len(digits) == 11 and digits.startswith("09"):
        return digits
    if len(digits) == 10 and digits.startswith("21"):
        return "0" + digits
    if len(digits) == 11 and digits.startswith("0"):
        return digits
    # keep original-ish text for ranges / extensions
    if "-" in s2 or "/" in s2:
        return s2 if s2.startswith("0") or not s2[0].isdigit() else s2
    # incomplete: keep digits as text without inventing prefix beyond mobile/21 rules
    return digits if len(digits) >= 7 else s2


def split_phones(cell) -> list[str]:
    if cell is None:
        return []
    s = str(cell).strip().translate(DIGITS)
    if not s:
        return []
    parts = re.split(r"[\n|;]+", s)
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        # also split "tel1 - tel2" when both look like numbers? keep ranges together if short
        ph = format_phone(p)
        if ph:
            out.append(ph)
        else:
            # try extracting multiple 09xxxxxxxxx
            for m in re.findall(r"0?9\d{9}", p.translate(DIGITS)):
                out.append(format_phone(m))
            for m in re.findall(r"0?21\d{8}", p.translate(DIGITS)):
                out.append(format_phone(m))
    # unique preserve order
    seen = set()
    uniq = []
    for x in out:
        if x and x not in seen:
            seen.add(x)
            uniq.append(x)
    return uniq


class Lead:
    __slots__ = ("company", "person", "phones", "sources", "uncertain", "title")

    def __init__(self, company, person=None, phones=None, source="", uncertain=True, title=""):
        self.company = norm_name(company)
        self.person = norm_name(person) if person else ""
        self.phones = list(phones or [])
        self.sources = {source} if source else set()
        self.uncertain = uncertain
        self.title = norm_name(title) if title else ""


def add_lead(bucket: dict, company, person=None, phones=None, source="", uncertain=True, title=""):
    company = norm_name(company)
    if not company or company in ("نام شرکت", "نام برند", "نام برند ", "ردیف"):
        return
    if company_key(company) in ("", "none", "nan"):
        return
    person = norm_name(person) if person else ""
    # drop garbage person values that are actually numbers
    if person and re.fullmatch(r"[\d\.\-]+", person.translate(DIGITS)):
        person = ""
    phones = [p for p in (phones or []) if p]
    ck = company_key(company)
    pk = person_key(person) if person else ""
    key = (ck, pk)
    if key not in bucket:
        bucket[key] = Lead(company, person, phones, source, uncertain, title)
        return
    lead = bucket[key]
    # prefer longer/cleaner company display
    if len(company) > len(lead.company):
        lead.company = company
    if title and not lead.title:
        lead.title = title
    if person and not lead.person:
        lead.person = person
    for p in phones:
        if p not in lead.phones:
            lead.phones.append(p)
    lead.sources.add(source)
    # if confirmed from multiple solid sources with phones, clear uncertain
    if not uncertain and phones:
        # keep uncertain false once confirmed
        if len(lead.sources) >= 2 and lead.phones:
            lead.uncertain = False
        else:
            lead.uncertain = lead.uncertain and uncertain
    elif uncertain is False:
        lead.uncertain = False


def read_workbook_full(bucket):
    path = FILES["workbook_full"]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        company, mobile, land, person, title = row[1], row[2], row[3], row[4], row[5]
        phones = split_phones(mobile) + split_phones(land)
        # solid if has person+phone or two phones
        unc = not (person and phones)
        add_lead(bucket, company, person, phones, "workbook_full", unc, title or "")
    wb.close()


def read_info_companies(bucket):
    path = FILES["info_companies"]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        company, region, phone, email, person = row[0], row[1], row[2], row[3], row[4]
        phones = split_phones(phone)
        # region may help display — append lightly if company bare
        company_disp = norm_name(company)
        if region and company_disp and str(region).strip() not in company_disp:
            # don't mutate name with region into company column permanently; keep company
            pass
        unc = True  # research queue = leads
        add_lead(bucket, company, person, phones, "info_companies", unc)
    wb.close()


def read_brand_sheets(bucket, path: Path, tag: str):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue
        header = [norm_name(c).replace("\u200c", " ") for c in rows[0]]
        # map columns
        def find(*cands):
            for i, h in enumerate(header):
                for c in cands:
                    if c in h.replace(" ", ""):
                        return i
            return None

        # brand sheets
        idx_brand = find("نامبرند", "نامشرکت") or (1 if "برند" in sheet or "گزارش" in sheet else 0)
        # heuristic by sheet type
        if "گزارش" in sheet:
            # ردیف، نام برند، ش جهت ارسال، ایمیل، شخص موثر، سمت
            for row in rows[1:]:
                if not row:
                    continue
                brand = row[1] if len(row) > 1 else None
                send = row[2] if len(row) > 2 else None
                person = row[4] if len(row) > 4 else None
                title = row[5] if len(row) > 5 else None
                phones = split_phones(send)
                add_lead(bucket, brand, person, phones, f"{tag}:{sheet}", True, title or "")
        elif "برند" in sheet or "همه" in sheet or sheet.strip().startswith("برند"):
            for row in rows[1:]:
                if not row or len(row) < 2:
                    continue
                brand = row[1]
                p1 = row[2] if len(row) > 2 else None
                p2 = row[3] if len(row) > 3 else None
                send = row[5] if len(row) > 5 else None
                phones = split_phones(p1) + split_phones(p2) + split_phones(send)
                add_lead(bucket, brand, None, phones, f"{tag}:{sheet}", True)
        # skip 1403 دهقان tracker sheet mostly — still harvest send names?
        elif "دهقان" in sheet:
            for row in rows[1:]:
                if not row:
                    continue
                # columns: brands ready, None, sent names, status...
                sent = row[2] if len(row) > 2 else None
                missing = row[5] if len(row) > 5 else None
                if sent and not str(sent).startswith("ارسال"):
                    add_lead(bucket, sent, None, [], f"{tag}:dehqan_sent", True)
                if missing and "راه" not in str(missing):
                    add_lead(bucket, missing, None, [], f"{tag}:dehqan_missing", True)
    wb.close()


def read_customer1399(bucket):
    path = FILES["customer1399"]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if not row:
                continue
            name = row[0]
            if not name:
                continue
            n = norm_name(name)
            if n in ("نام شرکت",) or n.startswith("برندهایی"):
                continue
            # brands without company sheet still as company-level lead
            add_lead(bucket, name, None, [], f"customer1399:{sheet}", True)
    wb.close()


def load_apir_members(bucket):
    # from git history
    try:
        raw = subprocess.check_output(
            [
                "git",
                "show",
                "origin/cursor/iran-distribution-b2b-6c45:research/iran-distribution-b2b/association_members_raw.json",
            ],
            cwd=r"D:\HomeDesign",
        )
        data = json.loads(raw.decode("utf-8"))
        (LOCAL / "association_members_raw.json").write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception as e:
        print("APIR json load fail", e)
        data = []
    for m in data:
        company = m.get("name_raw")
        person = m.get("ceo_source")
        phone = m.get("phone_official")
        phones = split_phones(phone)
        # official association list → less uncertain if phone present
        unc = not bool(phones)
        add_lead(bucket, company, person, phones, "apir_members", unc, "مدیرعامل/منبع انجمن")
    # APIR office contacts from public page
    add_lead(bucket, "انجمن ملی صنعت پخش ایران", "شایان مصری", ["02144009809"], "apir_contact", False, "مدیر اجرایی")
    add_lead(bucket, "انجمن ملی صنعت پخش ایران", "شهرزاد وزیری", ["02144009809"], "apir_contact", False, "مدیر دفتر")
    add_lead(bucket, "انجمن ملی صنعت پخش ایران", None, ["02144009809", "02144093720"], "apir_contact", False)
    # board as company-level person leads (no personal mobile on page)
    board = [
        ("دراژه", "سهراب کارگر", "رئیس هیئت مدیره انجمن / مدیرعامل"),
        ("هلدینگ شفا", "محمدرضا مرادی", "نائب رئیس انجمن / مدیرعامل"),
        ("قاسم ایران", "حسین یوسف خزرائی", "خزانه‌دار انجمن / مدیرعامل"),
        ("آرتا مهر", "مجید جعفر نژاد", "عضو هیئت مدیره انجمن / مدیرعامل"),
        ("طلوع پخش آفتاب", "حمیدرضا صفرشریف", "عضو هیئت مدیره انجمن"),
        ("سایه سمن", "مقصود نصیرزاده", "عضو هیئت مدیره انجمن / مدیرعامل"),
        ("گلپخش اول", "مجتبی کریمی", "عضو هیئت مدیره انجمن / مدیرعامل"),
        ("راسن درمان", "نادر علی زاد نمینی", "مدیرعامل"),
        ("الوند یاران سحر", "محمود توتونچیان", "علی‌البدل هیئت مدیره / مدیرعامل"),
        ("پراور پخش سپاهان", "حسین همتی نژاد", "علی‌البدل هیئت مدیره / مدیرعامل"),
    ]
    for co, pe, ti in board:
        add_lead(bucket, co, pe, [], "apir_board", True, ti)


def enrich_known_public(bucket):
    """Public official contact pages for priority distributors (confirmed when from company site)."""
    extras = [
        ("به پخش", None, ["02160912401", "02160912402"], "public_apir_member_phone", False),
        ("گلپخش اول", "مجتبی کریمی", ["02142418000", "02145416"], "public_gpa_ir", False, "مدیرعامل"),
        ("گلپخش اول", None, ["02142418000", "02145416"], "public_gpa_ir", False),
        ("سایه سمن", "مقصود نصیرزاده", ["02188800892"], "public_sayesaman", False, "مدیرعامل"),
        ("سایه سمن", None, ["02188800892", "02188800893", "02188800895"], "public_sayesaman", False),
        ("رضوان پخش عظیمی", None, ["90000405"], "public_rezvanpakhsh", False),
        ("رضوان پخش عظیمی", "خانم رنجبر", ["09023003042"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم بهزادی", ["09025005049", "09190079040"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "آقای سهرابی", ["09023003088"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم کریمی", ["09024300300", "09022000310"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "آقای رئوفی", ["09025400400"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم شمیرانی", ["09025005095"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم اختیاری", ["09023003083"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم قمرپور", ["09022000710"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم صادقی", ["09023003043"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم برغمدی", ["09021619000"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم موسی پور", ["09027300300"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم محقق", ["09022000975"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم چناقچی", ["09022000925"], "public_rezvanpakhsh", False, "کارشناس فروش"),
        ("رضوان پخش عظیمی", "خانم قنبری", ["09022000180"], "public_rezvanpakhsh", False, "کارشناس فروش"),
    ]
    for item in extras:
        if len(item) == 5:
            co, pe, ph, src, unc = item
            title = ""
        else:
            co, pe, ph, src, unc, title = item
        add_lead(bucket, co, pe, ph, src, unc, title)


def merge_person_display(lead: Lead) -> str:
    if lead.person and lead.title:
        # only append title if not already in person
        if lead.title not in lead.person:
            return f"{lead.person} — {lead.title}"
    return lead.person


def write_excel(leads: list[Lead], path: Path):
    # sort: companies with phones first, then alpha
    def sort_key(L: Lead):
        return (0 if L.phones else 1, company_key(L.company), person_key(L.person))

    leads = sorted(leads, key=sort_key)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "دفتر تماس"
    ws.sheet_view.rightToLeft = True
    headers = ["شرکت", "فرد مؤثر", "شماره موبایل و تلفن"]
    ws.append(headers)
    for col in range(1, 4):
        cell = ws.cell(1, col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", readingOrder=2)
        cell.border = THIN
    for L in leads:
        phones_text = "\n".join(L.phones) if L.phones else ""
        person = merge_person_display(L)
        ws.append([L.company, person, phones_text])
        r = ws.max_row
        for c in range(1, 4):
            cell = ws.cell(r, c)
            cell.font = BODY_FONT
            cell.alignment = Alignment(
                horizontal="right",
                vertical="center",
                wrap_text=True,
                readingOrder=2,
            )
            cell.border = THIN
            if c == 3:
                cell.number_format = "@"
                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center",
                    wrap_text=True,
                    readingOrder=1,
                )
            if L.uncertain:
                cell.fill = YELLOW
        # force phone as text
        ws.cell(r, 3).value = phones_text if phones_text else None
        ws.cell(r, 3).number_format = "@"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:C{ws.max_row}"
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 28
    ws.row_dimensions[1].height = 24
    for r in range(2, ws.max_row + 1):
        # taller if multi-line phones
        val = ws.cell(r, 3).value or ""
        lines = max(1, str(val).count("\n") + 1)
        ws.row_dimensions[r].height = max(18, 16 * lines)
    wb.save(path)
    return len(leads)


def main():
    bucket: dict[tuple, Lead] = {}
    print("reading workbook_full...")
    read_workbook_full(bucket)
    print(" reading info_companies...")
    read_info_companies(bucket)
    print(" reading brand sheets...")
    read_brand_sheets(bucket, FILES["brand_contact_a"], "brand_contact")
    read_brand_sheets(bucket, FILES["brand_b"], "brands")
    print(" reading customer1399...")
    read_customer1399(bucket)
    print(" loading APIR...")
    load_apir_members(bucket)
    enrich_known_public(bucket)

    leads = list(bucket.values())
    # drop empty / numeric-only garbage brand codes
    cleaned = []
    for L in leads:
        ck = company_key(L.company)
        if not ck:
            continue
        if re.fullmatch(r"\d{1,6}", ck):
            continue
        cleaned.append(L)
    leads = cleaned

    # stats
    with_phone = sum(1 for L in leads if L.phones)
    with_person = sum(1 for L in leads if L.person)
    uncertain = sum(1 for L in leads if L.uncertain)

    out = OUT_DIR / "Distribution-Contact-Directory-FA.xlsx"
    n = write_excel(leads, out)
    # also copy to research
    out2 = LOCAL / "Distribution-Contact-Directory-FA.xlsx"
    write_excel(leads, out2)

    meta = {
        "rows": n,
        "with_phone": with_phone,
        "with_person": with_person,
        "uncertain_yellow": uncertain,
        "sources": sorted({s for L in leads for s in L.sources}),
        "output": str(out),
    }
    (LOCAL / "build_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (OUT_DIR / "Distribution-Contact-Directory-build_meta.json").write_text(
        json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
