# -*- coding: utf-8 -*-
"""Import QUMS provincial distributor list + APIR national board CEOs + Daroopakhsh execs."""
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


def first_land(*vals):
    for v in vals:
        if not v:
            continue
        # take first phone-like token
        for tok in re.split(r"[\s,/،]+", str(v)):
            ld = land(tok)
            if ld:
                return ld
            # patterns like 024-33775603
            d = re.sub(r"\D", "", tok.translate(DIGITS))
            if d.startswith("98"):
                d = "0" + d[2:]
            ld = land(d)
            if ld:
                return ld
    return None


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


# From qums.ac.ir provincial external companies table (search snippet)
QUMS = [
    ("البرز — زنجان", "آقای غفاری", "02433775603"),
    ("داروپخش — تهران غرب", "آقای طیبی زاده", "02144545606"),
    ("محیا دارو", "آقای فولای", "02146883325"),
    ("رازی — زنجان", "آقای سخایی", "02432221147"),
    ("بهستان پخش", "آقای مهندس کرمی", "02188535311"),
    ("طب نوین", "آقای قملاقی", "02633215801"),
    ("به رسان دارو", "آقای معینی", "02432784992"),
    ("مهبان دارو", "آقای جعفری", "02146833360"),
    ("امید پخش آروند پارسیان", "آقای نقدیان", "02432221195"),
    ("لیوار", "آقای علیرضا صبحانی", "02636304330"),
    ("اکسیر زنجان", "آقای ساویز قربانی", "02432221242"),
    ("داروگستر باریج اسانس", "خانم امین صفار", "02140446341"),
    ("توسعه تجارت جلال آرا", "آقای ترابی", "02432221086"),
    ("هنار دارو", "خانم دکتر ملیحه مقیسه", "02144907334"),
    ("الیت دارو", "آقای کردی", "02636108908"),
    ("شفا آراد", "آقای شیرخانی", "02188504167"),
    ("یاقوت سلامت", "آقای قنبری", "02633524796"),
    ("پخش ممتاز — کرج", "آقای خانجانی", "02636670715"),
    ("دایا دارو", "آقای توانا", "02146048574"),
    ("پورا پخش", "آقای امینی", "02634763147"),
    ("سینا پخش صبا", "آقای نیک صالحی", "02636121052"),
    ("یارا طب ثامن", "آقای سلیمانی", "02144692133"),
    ("نفس نیک فردای سالم", "دکتر نظری", "02122581404"),
    ("به پخش بهشهر", "آقای یعقوبی", "02166258193"),
    ("سها هلال", "آقای امامی", "02177780198"),
]

# APIR national board with their companies
APIR_NAT = [
    ("سهراب کارگر", "رئیس هیئت مدیره انجمن ملی / مدیرعامل دراژه", "دراژه"),
    ("محمدرضا مرادی", "نایب رئیس انجمن ملی / مدیرعامل هلدینگ شفا", "هلدینگ شفا"),
    ("حسین یوسف خزرائی", "خزانه‌دار انجمن ملی / مدیرعامل قاسم ایران", "قاسم ایران"),
    ("مجید جعفر نژاد", "عضو هیئت مدیره انجمن ملی / مدیرعامل آرتا مهر", "آرتا مهر"),
    ("حمیدرضا صفرشریف", "عضو هیئت مدیره انجمن ملی / رئیس هیئت مدیره طلوع پخش آفتاب", "طلوع پخش آفتاب"),
    ("مقصود نصیرزاده", "عضو هیئت مدیره انجمن ملی / مدیرعامل سایه سمن", "سایه سمن"),
    ("مجتبی کریمی", "عضو هیئت مدیره انجمن ملی / مدیرعامل گلپخش اول", "گلپخش اول"),
    ("نادر علی زاد نمینی", "بازرس انجمن ملی / مدیرعامل راسن درمان", "راسن درمان"),
    ("احمد شیروانی", "دبیر کل انجمن ملی صنعت پخش", "انجمن ملی صنعت پخش ایران"),
    ("محمود توتونچیان", "عضو علی‌البدل انجمن ملی / مدیرعامل الوند یاران سحر", "الوند یاران سحر"),
    ("حسین همتی نژاد", "عضو علی‌البدل انجمن ملی / مدیرعامل پراور پخش سپاهان", "پراور پخش سپاهان"),
    ("سمیرا نصیری", "عضو علی‌البدل انجمن ملی", "انجمن ملی صنعت پخش ایران"),
    ("جبار صفی", "بازرس علی‌البدل انجمن ملی / مدیرعامل کاسپین", "کاسپین"),
]

DAROOPAKHSH = [
    ("علی مقدسی راد", "مدیرعامل و نایب رئیس هیئت مدیره توزیع داروپخش", "02188889600"),
    ("علی شریف اعلم", "رئیس هیئت مدیره توزیع داروپخش", "02188889600"),
    ("داودرضا شفاعت", "عضو هیئت مدیره توزیع داروپخش", "02188889400"),
    ("محمدرضا سحری", "عضو هیئت مدیره / معاون مالی توزیع داروپخش", "02188889400"),
    ("محمدحسین ملک", "عضو هیئت مدیره توزیع داروپخش", "02188889400"),
    ("پیمان عنبری", "CTO/CIO توزیع داروپخش", None),
    ("علیرضا خض", "CEO توزیع داروپخش (LinkedIn)", None),
    ("سامان مرادی", "مدیر مرکز داروپخش همدان/لرستان", None),
]


def try_fetch_qums():
    url = "https://qums.ac.ir/ar-SA/vcdf.qums.ac/7069/page/%D8%B4%D8%B1%DA%A9%D8%AA-%D9%87%D8%A7%DB%8C-%D8%AE%D8%A7%D8%B1%D8%AC-%D8%A7%D8%B3%D8%AA%D8%A7%D9%86%DB%8C"
    try:
        req = urllib.request.Request(url, headers=UA)
        html = urllib.request.urlopen(req, timeout=25).read().decode("utf-8", "ignore")
        (RAW / "qums_external_companies.html").write_text(html, encoding="utf-8")
        return html
    except Exception as e:
        print({"qums_fetch_error": str(e)})
        return ""


def parse_qums_html(html: str) -> list[dict]:
    hits = []
    if not html:
        return hits
    # table rows
    for m in re.finditer(
        r"<tr[^>]*>\s*<td[^>]*>\s*\d+\s*</td>\s*<td[^>]*>\s*([^<]+)</td>\s*<td[^>]*>\s*([^<]+)</td>\s*<td[^>]*>\s*([^<]*)</td>\s*<td[^>]*>\s*([^<]+)</td>",
        html,
        re.I | re.S,
    ):
        co = re.sub(r"\s+", " ", m.group(1)).strip()
        phone = re.sub(r"\s+", " ", m.group(2)).strip()
        pe = re.sub(r"\s+", " ", m.group(4)).strip()
        pe = pe.replace("آقای ", "").replace("خانم ", "").replace("مهندس ", "").replace("دکتر ", "")
        if not co or not pe or pe in ("مدیر شرکت",):
            continue
        hits.append(
            {
                "company": f"{co} — قزوین/البرز (فهرست علوم پزشکی)",
                "person": f"{pe} (مدیر شرکت — فهرست QUMS)",
                "mobile": to_mobile(phone),
                "land": first_land(phone),
                "src": "qums_table",
            }
        )
    return hits


def try_apir_members_file():
    """Try to find downloadable member list on APIR members page."""
    hits = []
    try:
        req = urllib.request.Request("https://www.apir.ir/pages/member-companies", headers=UA)
        html = urllib.request.urlopen(req, timeout=20).read().decode("utf-8", "ignore")
        (RAW / "apir_member_companies.html").write_text(html, encoding="utf-8")
        links = re.findall(r'href="([^"]+\.(?:xlsx?|pdf|csv)[^"]*)"', html, re.I)
        print({"apir_member_files": links})
        for link in links:
            url = link if link.startswith("http") else ("https://www.apir.ir" + link if link.startswith("/") else "https://www.apir.ir/" + link)
            try:
                req2 = urllib.request.Request(url, headers=UA)
                data = urllib.request.urlopen(req2, timeout=60).read()
                fname = RAW / ("apir_members_" + re.sub(r"\W+", "_", url)[-50:])
                if url.lower().endswith(".pdf"):
                    fname = fname.with_suffix(".pdf")
                elif "xls" in url.lower():
                    fname = fname.with_suffix(".xlsx")
                else:
                    fname = fname.with_suffix(".bin")
                fname.write_bytes(data)
                print({"saved": str(fname), "bytes": len(data)})
            except Exception as e:
                print({"member_dl_error": url, "err": str(e)})
    except Exception as e:
        print({"apir_members_error": str(e)})
    return hits


def main():
    hits: list[dict] = []
    html = try_fetch_qums()
    parsed = parse_qums_html(html)
    print({"qums_parsed": len(parsed)})
    hits.extend(parsed)

    for co, pe, phone in QUMS:
        pe2 = pe.replace("آقای ", "").replace("خانم ", "")
        hits.append(
            {
                "company": co,
                "person": f"{pe2} (مدیر شرکت — فهرست QUMS)",
                "mobile": to_mobile(phone),
                "land": first_land(phone),
                "src": "qums_curated",
            }
        )

    for pe, role, co in APIR_NAT:
        hits.append(
            {
                "company": co,
                "person": f"{pe} ({role})",
                "mobile": None,
                "land": land("02144009809"),
                "src": "apir_national_board",
            }
        )

    for pe, role, phone in DAROOPAKHSH:
        hits.append(
            {
                "company": "توزیع داروپخش",
                "person": f"{pe} ({role})",
                "mobile": None,
                "land": land(phone) if phone else None,
                "src": "daroopakhsh_exec",
            }
        )

    try_apir_members_file()
    added, wm = append_hits(hits)
    print(json.dumps({"prepared": len(hits), "appended": added, "with_mobile": wm}, ensure_ascii=False))


if __name__ == "__main__":
    main()
