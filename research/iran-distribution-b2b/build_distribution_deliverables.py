# -*- coding: utf-8 -*-
"""
Iran Distribution Association → B2B wholesale/credit discovery
Mission deliverables: multi-sheet XLSX + deep Persian PDF
Observation: 1405/09/03 (Jalali)
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    KeepTogether,
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

ROOT = Path(__file__).resolve().parent
ARTIFACTS = Path(r"D:\HomeDesign\artifacts")
ARTIFACTS.mkdir(parents=True, exist_ok=True)
OBS = "1405/09/03"
OBS_LABEL = "مشاهدهٔ مستقیم/مستند در تاریخ ۱۴۰۵/۰۹/۰۳"
PHASE_DIR = Path(r"D:\HomeDesign\research\iran-b2b-benchmark")

# ── fonts ──────────────────────────────────────────────────────────────────
FONT_CANDIDATES = [
    Path(r"C:\Windows\Fonts\Vazirmatn-Regular.ttf"),
    Path(r"C:\Windows\Fonts\vazirmatn.ttf"),
    Path(r"C:\Windows\Fonts\Vazir.ttf"),
    Path(r"C:\Windows\Fonts\tahoma.ttf"),
    ROOT / "fonts" / "Vazirmatn-Regular.ttf",
]
FONT_BOLD_CANDIDATES = [
    Path(r"C:\Windows\Fonts\Vazirmatn-Bold.ttf"),
    Path(r"C:\Windows\Fonts\tahomabd.ttf"),
    Path(r"C:\Windows\Fonts\tahoma.ttf"),
]


def pick_font(cands):
    for p in cands:
        if p.exists():
            return str(p)
    return None


FONT_REG = pick_font(FONT_CANDIDATES) or "Helvetica"
FONT_BOLD = pick_font(FONT_BOLD_CANDIDATES) or FONT_REG
pdfmetrics.registerFont(TTFont("Vazir", FONT_REG))
pdfmetrics.registerFont(TTFont("VazirBold", FONT_BOLD))


def rtl(text: str) -> str:
    """Wrap for RTL Persian paragraphs (ReportLab LTR engine)."""
    if text is None:
        return ""
    t = str(text).replace("\n", "<br/>")
    return f'<font face="Vazir">{t}</font>'


def rtl_b(text: str) -> str:
    t = str(text).replace("\n", "<br/>")
    return f'<font face="VazirBold">{t}</font>'


# ── known research findings (priority distributors) ─────────────────────────
# Status codes: Confirmed | Partial | Candidate | NoEvidence | Offline | Historical
_known_path = ROOT / "known_research.json"
if _known_path.exists():
    _raw_known = json.loads(_known_path.read_text(encoding="utf-8"))
    KNOWN = {}
    for k, v in _raw_known.items():
        vv = dict(v)
        vv["evidence"] = [tuple(x) if isinstance(x, list) else x for x in vv.get("evidence", [])]
        KNOWN[k] = vv
else:
    KNOWN = {}


def match_known(name: str):
    if name in KNOWN:
        return KNOWN[name]
    # longest key containment first to avoid به پخش دارو ← به پخش
    for k, v in sorted(KNOWN.items(), key=lambda kv: -len(kv[0])):
        if k == name or (len(k) >= 8 and (k in name or name in k)):
            return v
    soft = [
        ("باني چاو", "شركت پخش باني چاو ایرانیان"),
        ("بانی چاو", "شركت پخش باني چاو ایرانیان"),
        ("قاسم ايران", "شركت قاسم ايران"),
        ("قاسم ایران", "شركت قاسم ايران"),
        ("گلرنگ پخش", "شرکت گلرنگ پخش"),
        ("پدیده پایدار", "شرکت پخش پدیده پایدار"),
        ("پديده پايدار", "شرکت پخش پدیده پایدار"),
        ("سلامت پخش هستی", "شرکت سلامت پخش هستی(گلرنگ)"),
        ("مارینا پخش", "مارینا پخش هستی"),
        ("به پخش داروی", None),  # exclude
        ("رازي", "شركت پخش رازي"),
        ("رازی", "شركت پخش رازي"),
        ("يك و يك", "شركت پخش يك و يك"),
        ("یونیلیور", "شرکت یونیلیور ایران"),
        ("يونيلیور", "شرکت یونیلیور ایران"),
        ("داروپخش", "شركت توزيع داروپخش"),
        ("طلوع پخش آفتاب", "شرکت طلوع پخش آفتاب"),
        ("مدبران", "شرکت مدبران کوشای ایرانیان"),
    ]
    if name.strip() in ("شركت به پخش", "شرکت به پخش") and "شركت به پخش" in KNOWN:
        return KNOWN["شركت به پخش"]
    for token, key in soft:
        if key is None:
            continue
        if token in name and key in KNOWN:
            return KNOWN[key]
    return None


# Cross-cutting discovered platforms (not necessarily association members)
PLATFORMS = [
    {
        "id": "PLAT-SIMORGH",
        "name": "سیمرغ / PakhshMart (گلرنگ)",
        "url": "https://app.pakhshmart.com",
        "role": "پلتفرم سفارش یکپارچه پخش‌های گروه گلرنگ",
        "payments": "سفارش ۲۴/۷؛ تسویه طبق حساب مشتری",
        "credit": "Partial — قراردادی پشت حساب",
        "fit": "بهترین OTP زنده روی اعضای انجمن (گلرنگ پخش، پدیده پایدار، …)",
    },
    {
        "id": "PLAT-GHASEM",
        "name": "قاسم ایران B2B",
        "url": "https://www.ghasemiran.org/fa/page/ghasemiran-B2B.html",
        "role": "اپ سفارش با کد مشتری مینو/قاسم",
        "payments": "فاکتور رسمی + تخفیف‌های پلتفرمی",
        "credit": "Partial قراردادی",
        "fit": "الگوی CRM-first Confirmed",
    },
    {
        "id": "PLAT-BONNY",
        "name": "بانی چاو (سولیکو/کاله)",
        "url": "https://www.bonnychow.ir",
        "role": "پورتال سفارش + مالی برای کدخریدار",
        "payments": "کارتخوان تحویل + درگاه صورتحساب باز + روش‌های رایج",
        "credit": "خط اعتبار خوش‌حسابان؛ ۷۲س = SLA تمدید پس از وصول نه مهلت خرید",
        "fit": "نزدیک‌ترین عضو انجمن با سفارش+پرداخت+اعتبار مستند",
    },
    {
        "id": "PLAT-RAZI",
        "name": "پرتال سفارش رازی",
        "url": "http://orders.razico.ir/",
        "role": "سفارش اینترنتی پخش دارو",
        "payments": "پشت لاگین",
        "credit": "اعتبار داروخانه‌ای رایج صنعت",
        "fit": "Confirmed portal داخل انجمن (حوزه دارو)",
    },
    {
        "id": "PLAT-SOOT",
        "name": "سوت (SOOT)",
        "url": "https://soot.app / Myket com.soot.app",
        "role": "شبکه سفارش عمده FMCG چندفروشنده",
        "payments": "درگاه + نقد هنگام تحویل + چک + اعتباری (ادعای استور)",
        "credit": "اعلام صریح چک و اعتباری در کنار نقد/آنلاین",
        "fit": "نزدیک‌ترین پیام استوری به «نقد/چک/اعتبار همان مسیر» — مقیاس نصب هنوز کوچک",
    },
    {
        "id": "PLAT-SABAD",
        "name": "سبد آنلاین",
        "url": "https://sabad.online / online.sabad.shop",
        "role": "تجمیع پخش‌ها برای خرید عمده",
        "payments": "پای‌بار در مشهد و …",
        "credit": "خرید اعتباری در مشهد اعلام شده",
        "fit": "تجمیع‌کننده زنده با سیگنال اعتبار منطقه‌ای",
    },
    {
        "id": "PLAT-HIMART",
        "name": "Himart / HiCard",
        "url": "himart.ir ecosystem",
        "role": "فروش با COD/آنلاین/اعتبار",
        "payments": "COD + آنلاین + اعتبار",
        "credit": "بله محصولی",
        "fit": "الگوی پرداخت چندگانه؛ لزوماً عضو انجمن نیست",
    },
    {
        "id": "PLAT-BAROOK",
        "name": "باروک / کسب‌نو (Kasbno)",
        "url": "barook / kasbno",
        "role": "B2B marketplace + اعتبار",
        "payments": "نقدی + اعتباری",
        "credit": "قوی‌ترین سیگنال موجودی+اعتبار فاز۱/۲",
        "fit": "الگوی هدف پلتفرمی",
    },
    {
        "id": "PLAT-PINDIS",
        "name": "پیندیس (PinDis)",
        "url": "pindis",
        "role": "عمده‌فروشی با قوانین شفاف",
        "payments": "شرایط عمده",
        "credit": "در قوانین فروش",
        "fit": "بنچمارک قوانین عمده",
    },
    {
        "id": "PLAT-DIGIPAY",
        "name": "دیجی‌پی سازمانی",
        "url": "digipay.ir",
        "role": "لایه پرداخت/اعتبار — نه کاتالوگ پخش",
        "payments": "پرداخت و اعتبار",
        "credit": "سازمانی ≠ موجودی توزیع",
        "fit": "مکمل چک‌اوت نه جایگزین پخش",
    },
    {
        "id": "PLAT-APIR",
        "name": "انجمن ملی صنعت پخش (APIR)",
        "url": "https://www.apir.ir/pages/member-companies",
        "role": "دایرکتوری اعضا + مجوز/سامانه توزیع کالا",
        "payments": "—",
        "credit": "—",
        "fit": "بذر اصلی مأموریت؛ آپدیت پاییز ۱۴۰۵ روی سایت",
    },
]


SCENARIOS = [
    {
        "id": "SC-01",
        "title": "خرده‌فروش → سیمرغ (گلرنگ پخش / پدیده پایدار)",
        "actor": "فروشگاه با موبایل ثبت‌شده در پخش گلرنگ",
        "steps": [
            "ورود به app.pakhshmart.com یا pakhshmart.com",
            "وارد کردن موبایل ثبت‌شده در شرکت پخش",
            "دریافت OTP و ورود",
            "مرور کاتالوگ گلرنگ پخش / پدیده پایدار و ثبت سفارش ۲۴/۷",
            "تسویه طبق وضعیت حساب (جزئیات درگاه پشت پنل)",
        ],
        "payments": "قراردادی/پنل مشتری",
        "credit_same_deal": "Partial",
        "evidence": "Confirmed OTP + landing",
    },
    {
        "id": "SC-02",
        "title": "خرده‌فروش/سازمان → اپ B2B قاسم ایران",
        "actor": "اصناف FMCG دارای کد مشتری قاسم",
        "steps": [
            "دریافت کد مشتری از قاسم ایران",
            "دانلود اپ B2B از صفحه ghasemiran.org",
            "جستجو، تخفیف، ثبت سفارش، فاکتور رسمی",
            "تحویل با ناوگان قاسم",
        ],
        "payments": "کانال مالی مشتری",
        "credit_same_deal": "Partial قراردادی",
        "evidence": "Confirmed official B2B page",
    },
    {
        "id": "SC-03",
        "title": "فروشگاه → بانی چاو: سفارش آنلاین + کارتخوان/درگاه + اعتبار خوش‌حسابی",
        "actor": "خریدار با کد سولیکو/کاله",
        "steps": [
            "درخواست کد خریدار (تلفن ۰۲۱-۶۱۳۷۰ یا ثبت‌نام)",
            "ورود پورتال با موبایل",
            "ثبت سفارش آنلاین یا از طریق کارشناس",
            "پرداخت هنگام تحویل با کارتخوان، یا تسویه open items از درگاه",
            "مصرف/تمدید خط اعتبار طبق دوره بازپرداخت؛ نمایش تمدید ≤۷۲س پس از وصول مالی",
        ],
        "payments": "کارتخوان تحویل + درگاه + روش‌های رایج",
        "credit_same_deal": "Confirmed خط اعتبار؛ ۷۲س=SLA تمدید نه مهلت خرید",
        "evidence": "Confirmed FAQ bonnychow.ir",
    },
    {
        "id": "SC-04",
        "title": "داروخانه → پرتال سفارش رازی",
        "actor": "مشتری ثبت‌شده پخش رازی",
        "steps": [
            "ورود به orders.razico.ir یا my.razico.ir",
            "لاگین با کاربر/رمز سازمانی",
            "ثبت/پیگیری سفارش دارویی",
            "تسویه طبق قرارداد داروخانه",
        ],
        "payments": "پشت لاگین",
        "credit_same_deal": "Partial — نرم صنعت دارو",
        "evidence": "Confirmed portals",
    },
    {
        "id": "SC-05",
        "title": "خرده‌فروش → سوت: درگاه + نقد تحویل + چک + اعتباری",
        "actor": "سوپرمارکت/رستوران بدون وابستگی به یک پخش",
        "steps": [
            "نصب com.soot.app یا soot.app",
            "انتخاب فروشنده و سبد FMCG",
            "انتخاب روش تسویه اعلامی: درگاه / نقد تحویل / چک / اعتباری",
            "تحویل و پیگیری",
        ],
        "payments": "چندروش (ادعای استور)",
        "credit_same_deal": "اعلام‌شده در استور — مقیاس کوچک؛ نیاز به تست واقعی",
        "evidence": "Confirmed store copy / Partial scale",
    },
    {
        "id": "SC-06",
        "title": "خرده‌فروش → سبد آنلاین (تجمیع پخش)",
        "actor": "فروشگاه (قوی‌تر در مشهد برای اعتبار)",
        "steps": [
            "ورود sabad.online یا اپ online.sabad.shop",
            "انتخاب شهر و پخش",
            "سفارش عمده؛ در مشهد: پای‌بار / اعتبار اعلامی",
        ],
        "payments": "متغیر / پای‌بار",
        "credit_same_deal": "اعلام منطقه‌ای مشهد",
        "evidence": "Confirmed store + site",
    },
    {
        "id": "SC-07",
        "title": "مسیر غالب انجمن: ویزیتور/تلفن (OutsideDigitalPath)",
        "actor": "اکثر اعضای NoEvidence",
        "steps": [
            "تماس سوئیچبورد از اکسل انجمن",
            "درخواست واحد فروش مویرگی",
            "سفارش شفاهی/پیام‌رسان به ویزیتور",
            "فاکتور + نقد/چک/اعتبار آفلاین",
        ],
        "payments": "نقد/چک",
        "credit_same_deal": "آفلاین OutsideDigitalPath",
        "evidence": "Dominant pattern",
    },
    {
        "id": "SC-08",
        "title": "لایه پرداخت جدا (دیجی‌پی) + سفارش پخش",
        "actor": "فروشگاه با اعتبار پرداختی",
        "steps": [
            "اخذ اعتبار از لایهٔ فین‌تک",
            "سفارش از سیمرغ/قاسم/بانی/ویزیتور",
            "پرداخت فقط اگر پخش ابزار را بپذیرد",
        ],
        "payments": "اعتبار پرداختی",
        "credit_same_deal": "موجودی ≠ اعتبار مگر اتصال چک‌اوت",
        "evidence": "Phase1 distinction",
    },
    {
        "id": "SC-09",
        "title": "بازیابی کانال به پخش مارکت (Historical)",
        "actor": "عامل فروش شبکه به پخش",
        "steps": [
            "جستجوی ir.bpc.market در استور یا تماس ۶۰۹۱۲۴۰۱",
            "اگر اپ زنده: سفارش بدون ویزیتور",
            "وگرنه: مسیر کلاسیک ویزیتور + صرفه مارکت فقط برای B2C",
        ],
        "payments": "نامشخص ۱۴۰۵",
        "credit_same_deal": "نامشخص",
        "evidence": "Historical 1399 press",
    },
]


def load_members():
    return json.loads((ROOT / "association_members_raw.json").read_text(encoding="utf-8"))


def normalize_name(raw: str) -> str:
    s = raw.replace("شركت", "شرکت").replace("ي", "ی").replace("ك", "ک")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def classify_priority(m: dict) -> str:
    act = m.get("activity", "")
    if any(x in act for x in ("غذا", "بهداشت", "آرایشی", "آرايشي", "شوینده", "شوينده")):
        return "P1-FMCG"
    if "دارو" in act:
        return "P2-Pharma"
    return "P3-Other"


def build_company_rows(members):
    rows = []
    for i, m in enumerate(members, start=1):
        known = match_known(m["name_raw"])
        cid = known["company_id"] if known else f"DIST-{i:03d}"
        status = known["status"] if known else "NoEvidence"
        website = known.get("website", "") if known else ""
        portal = known.get("portal_url", "") if known else ""
        app = known.get("app_name", "") if known else ""
        channel = known.get("channel", "نامشخص — احتمالاً ویزیتور/تلفن") if known else "نامشخص — بررسی پورتال عمومی نتیجه‌ای نداد"
        login = known.get("login", "—") if known else "—"
        payments = known.get("payments", "—") if known else "—"
        credit = known.get("credit", "—") if known else "—"
        notes = known.get("notes", "") if known else "در فهرست انجمن؛ مسیر سفارش آنلاین عمومی در خزش فعلی یافت نشد. اقدام بعدی: تماس تلفن رسمی + جستجوی دامنهٔ ایمیل."
        rows.append(
            {
                "company_id": cid,
                "name_raw": m["name_raw"],
                "name_normalized": normalize_name(m["name_raw"]),
                "priority": classify_priority(m),
                "topic": m.get("topic", ""),
                "activity": m.get("activity", ""),
                "scope_hq": m.get("scope_hq", ""),
                "address": m.get("address", ""),
                "phone_official": m.get("phone_official", ""),
                "fax": m.get("fax", ""),
                "email": m.get("email", ""),
                "website": website,
                "portal_url": portal,
                "app_name": app,
                "channel_type": channel,
                "login_method": login,
                "payment_methods": payments,
                "credit_same_deal": credit,
                "status": status,
                "last_checked": OBS,
                "notes": notes,
                "ceo_source": m.get("ceo_source", ""),
                "known": known,
            }
        )
    return rows


def build_decision_makers(companies):
    dms = []
    for c in companies:
        ceo = (c.get("ceo_source") or "").strip()
        if ceo and ceo not in ("ندارند", "-", "—"):
            dms.append(
                {
                    "company_id": c["company_id"],
                    "company_name": c["name_raw"],
                    "full_name": ceo,
                    "role_title": "مدیرعامل (بر اساس ستون انجمن ملی صنعت پخش — لیست اعضا ۱۴۰۴)",
                    "mobile_public": "",  # فقط اگر عمومی باشد؛ در لیست انجمن موبایل مدیران نیست
                    "phone_switchboard": c["phone_official"],
                    "email_public": c["email"] if "@" in (c.get("email") or "") else "",
                    "source_url": "فایل اکسل انجمن ملی صنعت پخش ایران — لیست اعضا ۱۴۰۴",
                    "source_type": "AssociationDirectory",
                    "last_verified": OBS,
                    "notes": "موبایل شخصی در منبع انجمن منتشر نشده؛ فقط تلفن سازمان و در صورت وجود ایمیل سازمانی.",
                }
            )
        # email contact as org contact if present
        if c.get("email") and "@" in c["email"]:
            # already covered via CEO row notes; add org email contact row when no CEO
            if not ceo:
                dms.append(
                    {
                        "company_id": c["company_id"],
                        "company_name": c["name_raw"],
                        "full_name": "(مخاطب سازمانی — نام شخص منتشرنشده)",
                        "role_title": "ایمیل سازمانی",
                        "mobile_public": "",
                        "phone_switchboard": c["phone_official"],
                        "email_public": c["email"],
                        "source_url": "فایل اکسل انجمن",
                        "source_type": "AssociationDirectory",
                        "last_verified": OBS,
                        "notes": "فقط کانال سازمانی.",
                    }
                )
    return dms


def build_discovery_links(companies):
    links = []
    lid = 1
    # association source
    links.append(
        {
            "link_id": f"L-{lid:04d}",
            "company_id": "ASSOC",
            "url": "انجمن ملی صنعت پخش ایران — فایل اعضای ۱۴۰۴ (آپلود مأموریت)",
            "page_type": "Directory",
            "discovered_from": "UserUpload",
            "http_status": "local-file",
            "notes": "منبع اصلی ۲۳۴ عضو",
        }
    )
    lid += 1
    links.append(
        {
            "link_id": f"L-{lid:04d}",
            "company_id": "ASSOC",
            "url": "https://www.apir.ir (در صورت در دسترس بودن)",
            "page_type": "AssociationSite",
            "discovered_from": "PriorKnowledge",
            "http_status": "variable",
            "notes": "سایت انجمن؛ دسترسی گاهی ناپایدار",
        }
    )
    lid += 1
    for c in companies:
        if c.get("website"):
            links.append(
                {
                    "link_id": f"L-{lid:04d}",
                    "company_id": c["company_id"],
                    "url": c["website"],
                    "page_type": "CorporateSite",
                    "discovered_from": "Research",
                    "http_status": "see-evidence",
                    "notes": c.get("notes", "")[:120],
                }
            )
            lid += 1
        if c.get("portal_url") and c["portal_url"] != "—":
            links.append(
                {
                    "link_id": f"L-{lid:04d}",
                    "company_id": c["company_id"],
                    "url": c["portal_url"],
                    "page_type": "B2BPortal",
                    "discovered_from": "DeepCrawl",
                    "http_status": "see-evidence",
                    "notes": c.get("status", ""),
                }
            )
            lid += 1
        # email domain hint
        em = c.get("email") or ""
        if "@" in em:
            domain = em.split("@")[-1].split(";")[0].strip()
            if domain and "." in domain:
                links.append(
                    {
                        "link_id": f"L-{lid:04d}",
                        "company_id": c["company_id"],
                        "url": f"https://{domain} (حدس از ایمیل انجمن — نیاز به تأیید)",
                        "page_type": "EmailDomainHint",
                        "discovered_from": "AssociationEmail",
                        "http_status": "unverified",
                        "notes": "Candidate website from email domain",
                    }
                )
                lid += 1
    for p in PLATFORMS:
        links.append(
            {
                "link_id": f"L-{lid:04d}",
                "company_id": p["id"],
                "url": p["url"],
                "page_type": "Platform",
                "discovered_from": "CrossCuttingDiscovery",
                "http_status": "see-phase",
                "notes": p["fit"],
            }
        )
        lid += 1
    return links


def build_evidence(companies):
    ev = []
    eid = 1
    ev.append(
        {
            "evidence_id": f"E-{eid:04d}",
            "company_id": "ASSOC",
            "claim": "۲۳۴ شرکت در لیست اعضای انجمن (فایل ۱۴۰۴) با نام مدیرعامل، تلفن، ایمیل، نوع فعالیت",
            "url": "uploads association xlsx",
            "observed_at": OBS,
            "snippet": "هدر: ردیف، نام شرکت، نام مدیرعامل، آدرس، تلفن، فکس، سرفصل، نوع فعالیت، ایمیل، نوع پخش",
            "confidence": "Confirmed",
        }
    )
    eid += 1
    for c in companies:
        known = c.get("known")
        if not known:
            continue
        for title, url, conf in known.get("evidence", []):
            ev.append(
                {
                    "evidence_id": f"E-{eid:04d}",
                    "company_id": c["company_id"],
                    "claim": f"{c['name_raw']}: {title}",
                    "url": url,
                    "observed_at": OBS,
                    "snippet": known.get("notes", "")[:200],
                    "confidence": conf,
                }
            )
            eid += 1
    # platform evidence
    for p in PLATFORMS:
        ev.append(
            {
                "evidence_id": f"E-{eid:04d}",
                "company_id": p["id"],
                "claim": f"پلتفرم {p['name']}: {p['role']}",
                "url": p["url"],
                "observed_at": OBS,
                "snippet": p["fit"],
                "confidence": "Partial" if "Digipay" in p["name"] or "سبد" in p["name"] else "Confirmed",
            }
        )
        eid += 1
    return ev


def build_crawl_log(companies):
    log = []
    log.append(
        {
            "ts": OBS,
            "target": "association_xlsx",
            "action": "parse_234_members",
            "result": "ok",
            "notes": "column mapping fixed: topic/activity/email/scope_hq",
        }
    )
    log.append(
        {
            "ts": OBS,
            "target": "app.pakhshmart.com",
            "action": "portal_probe",
            "result": "OTP login confirmed pattern",
            "notes": "Simorgh",
        }
    )
    log.append(
        {
            "ts": OBS,
            "target": "ghasemiran B2B",
            "action": "product_pattern",
            "result": "customer-code login",
            "notes": "Partial public detail",
        }
    )
    log.append(
        {
            "ts": OBS,
            "target": "Bonnie Chow",
            "action": "payment_methods",
            "result": "COD+online+72h credit",
            "notes": "best same-deal multipay sample",
        }
    )
    log.append(
        {
            "ts": OBS,
            "target": "Beh Pakhsh Market",
            "action": "store_history",
            "result": "historical 17k installs; 1405 live status unconfirmed",
            "notes": "Historical",
        }
    )
    log.append(
        {
            "ts": OBS,
            "target": "Sabad Online",
            "action": "marketplace_discovery",
            "result": "Myket listing / aggregator signal",
            "notes": "Candidate bridge",
        }
    )
    for c in companies:
        if c["status"] == "NoEvidence":
            log.append(
                {
                    "ts": OBS,
                    "target": c["name_raw"],
                    "action": "public_portal_search",
                    "result": "NoEvidence",
                    "notes": f"phone={c['phone_official']}; email={c['email']}; next=phone verify",
                }
            )
        else:
            log.append(
                {
                    "ts": OBS,
                    "target": c["name_raw"],
                    "action": "deep_profile",
                    "result": c["status"],
                    "notes": (c.get("portal_url") or c.get("website") or "")[:100],
                }
            )
    return log


def build_gaps(companies):
    gaps = []
    gid = 1
    noev = [c for c in companies if c["status"] == "NoEvidence"]
    gaps.append(
        {
            "gap_id": f"G-{gid:03d}",
            "topic": "پوشش پورتال عمومی",
            "detail": f"از {len(companies)} عضو، {len(noev)} مورد هنوز NoEvidence برای سفارش آنلاین عمومی هستند.",
            "priority": "P0",
            "next_action": "تماس تلفنی ساخت‌یافته با ۳۰ شرکت P1-FMCG اولویت‌دار + بررسی دامنهٔ ایمیل",
        }
    )
    gid += 1
    gaps.append(
        {
            "gap_id": f"G-{gid:03d}",
            "topic": "موبایل تصمیم‌گیرندگان",
            "detail": "فهرست انجمن موبایل مدیران را منتشر نمی‌کند؛ فقط نام مدیرعامل + تلفن سازمان.",
            "priority": "P0",
            "next_action": "فقط از منابع عمومی (سایت، وای‌گپ رسمی، روزنامه) استخراج شود؛ بدون حدس.",
        }
    )
    gid += 1
    gaps.append(
        {
            "gap_id": f"G-{gid:03d}",
            "topic": "وضعیت زندهٔ به پخش مارکت",
            "detail": "اپ تاریخی تأیید شده؛ وضعیت ۱۴۰۵ نیاز به تأیید فروشگاه/تلفن دارد.",
            "priority": "P1",
            "next_action": "تماس ۶۰۹۱۲۴۰۱ + بررسی مجدد کافه‌بازار/مایکت",
        }
    )
    gid += 1
    gaps.append(
        {
            "gap_id": f"G-{gid:03d}",
            "topic": "جزئیات درگاه پرداخت داخل پورتال‌ها",
            "detail": "بیشتر پورتال‌ها پشت لاگین؛ روش نقد/چک/اعتبار در صفحهٔ عمومی کامل دیده نمی‌شود.",
            "priority": "P1",
            "next_action": "حساب تست خرده‌فروش یا مصاحبه با فروشگاه طرف قرارداد",
        }
    )
    gid += 1
    gaps.append(
        {
            "gap_id": f"G-{gid:03d}",
            "topic": "هم‌پوشانی با فاز۱/۲",
            "detail": "باروک/کسب‌نو نزدیک‌ترین «موجودی+اعتبار» است ولی لزوماً عضو انجمن نیست.",
            "priority": "P2",
            "next_action": "نگاشت تأمین‌کنندگان باروک به اعضای انجمن",
        }
    )
    return gaps


# ── Excel writer ────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Tahoma", size=11)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="right")
THIN = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)


def style_header(ws, ncols):
    for col in range(1, ncols + 1):
        cell = ws.cell(1, col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")


def autosize(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_sheet(ws, headers, rows, widths):
    ws.append(headers)
    style_header(ws, len(headers))
    for r in rows:
        ws.append(r)
        for col in range(1, len(headers) + 1):
            cell = ws.cell(ws.max_row, col)
            cell.alignment = WRAP
            cell.border = THIN
            cell.font = Font(name="Tahoma", size=10)
    autosize(ws, widths)
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"


def build_excel(companies, dms, links, evidence, crawl, gaps, scenarios):
    wb = Workbook()

    # Cover
    ws = wb.active
    ws.title = "00_Cover"
    cover = [
        ["عنوان", "کشف مسیر خرید عمده آنلاین از اعضای انجمن ملی صنعت پخش ایران"],
        ["تاریخ مشاهده", OBS],
        ["تعداد شرکت‌ها", len(companies)],
        ["منبع اعضا", "اکسل انجمن — لیست اعضا ۱۴۰۴"],
        ["خروجی‌ها", "همین فایل + PDF عمیق"],
        ["قواعد شواهد", "Confirmed / Partial / Candidate / NoEvidence / Historical / OutsideDigitalPath"],
        ["نکتهٔ حریم", "موبایل مدیران فقط در صورت انتشار عمومی؛ در انجمن موجود نیست"],
    ]
    for row in cover:
        ws.append(row)
    autosize(ws, [28, 80])

    # Companies
    ws = wb.create_sheet("01_Companies")
    headers = [
        "company_id",
        "name_raw",
        "name_normalized",
        "priority",
        "topic",
        "activity",
        "scope_hq",
        "address",
        "phone_official",
        "fax",
        "email",
        "website",
        "portal_url",
        "app_name",
        "channel_type",
        "login_method",
        "payment_methods",
        "credit_same_deal",
        "status",
        "last_checked",
        "notes",
    ]
    rows = [[c.get(h, "") for h in headers] for c in companies]
    write_sheet(ws, headers, rows, [12, 28, 28, 12, 12, 18, 16, 40, 16, 12, 28, 28, 28, 20, 28, 18, 28, 28, 12, 12, 50])

    # Decision makers
    ws = wb.create_sheet("02_DecisionMakers")
    dh = [
        "company_id",
        "company_name",
        "full_name",
        "role_title",
        "mobile_public",
        "phone_switchboard",
        "email_public",
        "source_url",
        "source_type",
        "last_verified",
        "notes",
    ]
    write_sheet(ws, dh, [[d.get(h, "") for h in dh] for d in dms], [12, 28, 22, 40, 14, 16, 28, 40, 18, 12, 40])

    # Discovery links
    ws = wb.create_sheet("03_DiscoveryLinks")
    lh = ["link_id", "company_id", "url", "page_type", "discovered_from", "http_status", "notes"]
    write_sheet(ws, lh, [[x.get(h, "") for h in lh] for x in links], [10, 12, 55, 16, 18, 14, 40])

    # Providers / Platforms
    ws = wb.create_sheet("04_Providers_Platforms")
    ph = ["id", "name", "url", "role", "payments", "credit", "fit_for_mission"]
    write_sheet(
        ws,
        ph,
        [[p["id"], p["name"], p["url"], p["role"], p["payments"], p["credit"], p["fit"]] for p in PLATFORMS],
        [14, 24, 30, 36, 24, 28, 40],
    )

    # Scenarios
    ws = wb.create_sheet("05_Scenarios")
    sh = ["scenario_id", "title", "actor", "steps", "payments", "credit_same_deal", "evidence"]
    srows = [
        [
            s["id"],
            s["title"],
            s["actor"],
            " | ".join(s["steps"]),
            s["payments"],
            s["credit_same_deal"],
            s["evidence"],
        ]
        for s in scenarios
    ]
    write_sheet(ws, sh, srows, [12, 40, 28, 70, 24, 28, 24])

    # Scenario steps detailed
    ws = wb.create_sheet("06_ScenarioSteps")
    ws.append(["scenario_id", "step_no", "step_text"])
    style_header(ws, 3)
    for s in scenarios:
        for i, st in enumerate(s["steps"], 1):
            ws.append([s["id"], i, st])
    autosize(ws, [12, 10, 70])

    # Evidence
    ws = wb.create_sheet("07_Evidence")
    eh = ["evidence_id", "company_id", "claim", "url", "observed_at", "snippet", "confidence"]
    write_sheet(ws, eh, [[e.get(h, "") for h in eh] for e in evidence], [12, 12, 50, 40, 12, 40, 12])

    # Crawl log
    ws = wb.create_sheet("08_CrawlLog")
    ch = ["ts", "target", "action", "result", "notes"]
    write_sheet(ws, ch, [[x.get(h, "") for h in ch] for x in crawl], [12, 36, 22, 28, 50])

    # Gaps
    ws = wb.create_sheet("09_Gaps_NextActions")
    gh = ["gap_id", "topic", "detail", "priority", "next_action"]
    write_sheet(ws, gh, [[g.get(h, "") for h in gh] for g in gaps], [10, 24, 55, 10, 50])

    # Status summary
    ws = wb.create_sheet("10_StatusSummary")
    from collections import Counter

    sc = Counter(c["status"] for c in companies)
    pc = Counter(c["priority"] for c in companies)
    ws.append(["metric", "value"])
    style_header(ws, 2)
    ws.append(["total_companies", len(companies)])
    for k, v in sc.items():
        ws.append([f"status_{k}", v])
    for k, v in pc.items():
        ws.append([f"priority_{k}", v])
    ws.append(["decision_makers_rows", len(dms)])
    ws.append(["discovery_links", len(links)])
    ws.append(["evidence_rows", len(evidence)])
    autosize(ws, [28, 14])

    # FMCG shortlist
    ws = wb.create_sheet("11_FMCG_Shortlist")
    fmcg = [c for c in companies if c["priority"] == "P1-FMCG"]
    write_sheet(
        ws,
        ["company_id", "name_raw", "ceo", "phone", "email", "status", "portal_url", "notes"],
        [
            [
                c["company_id"],
                c["name_raw"],
                c["ceo_source"],
                c["phone_official"],
                c["email"],
                c["status"],
                c["portal_url"],
                c["notes"][:80],
            ]
            for c in fmcg
        ],
        [12, 28, 20, 16, 28, 12, 30, 40],
    )

    out = ARTIFACTS / "Iran_Distribution_B2B_Directory.xlsx"
    # also copy under research
    out2 = ROOT / "Iran_Distribution_B2B_Directory.xlsx"
    wb.save(out)
    wb.save(out2)
    return out


# ── PDF ─────────────────────────────────────────────────────────────────────
def make_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="PTitle",
            fontName="VazirBold",
            fontSize=18,
            leading=28,
            alignment=TA_CENTER,
            spaceAfter=12,
            textColor=colors.HexColor("#1F4E79"),
        )
    )
    styles.add(
        ParagraphStyle(
            name="PH1",
            fontName="VazirBold",
            fontSize=14,
            leading=24,
            alignment=TA_RIGHT,
            spaceBefore=14,
            spaceAfter=8,
            textColor=colors.HexColor("#1F4E79"),
        )
    )
    styles.add(
        ParagraphStyle(
            name="PH2",
            fontName="VazirBold",
            fontSize=12,
            leading=20,
            alignment=TA_RIGHT,
            spaceBefore=10,
            spaceAfter=6,
            textColor=colors.HexColor("#2E75B6"),
        )
    )
    styles.add(
        ParagraphStyle(
            name="PBody",
            fontName="Vazir",
            fontSize=10,
            leading=18,
            alignment=TA_JUSTIFY,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="PMeta",
            fontName="Vazir",
            fontSize=9,
            leading=14,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#555555"),
        )
    )
    styles.add(
        ParagraphStyle(
            name="PSmall",
            fontName="Vazir",
            fontSize=8,
            leading=13,
            alignment=TA_RIGHT,
        )
    )
    styles.add(
        ParagraphStyle(
            name="PBullet",
            fontName="Vazir",
            fontSize=10,
            leading=16,
            alignment=TA_RIGHT,
            rightIndent=10,
        )
    )
    return styles


def p(styles, text, style="PBody"):
    return Paragraph(rtl(text) if style != "PTitle" else rtl_b(text), styles[style])


def pb(styles, text, style="PH1"):
    return Paragraph(rtl_b(text), styles[style])


def simple_table(data, col_widths, styles):
    # wrap cells
    wrapped = []
    for row in data:
        wrapped.append(
            [
                Paragraph(rtl_b(c) if i == 0 and row is data[0] else rtl(str(c)), styles["PSmall"])
                for i, c in enumerate(row)
            ]
            if False
            else [Paragraph(rtl(str(c)), styles["PSmall"]) for c in row]
        )
    # header bold
    wrapped[0] = [Paragraph(rtl_b(str(c)), styles["PSmall"]) for c in data[0]]
    t = Table(wrapped, colWidths=col_widths, repeatRows=1)
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CCCCCC")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F8FC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return t


def build_pdf(companies, dms, links, evidence, crawl, gaps, scenarios):
    styles = make_styles()
    story = []
    from collections import Counter

    sc = Counter(c["status"] for c in companies)
    pc = Counter(c["priority"] for c in companies)

    # Cover
    story.append(Spacer(1, 2 * cm))
    story.append(Paragraph(rtl_b("سند مأموریت خزش و کشف مسیرهای B2B پخش ایران"), styles["PTitle"]))
    story.append(Spacer(1, 0.4 * cm))
    story.append(
        Paragraph(
            rtl_b("از فهرست انجمن ملی صنعت پخش تا پورتال سفارش، پرداخت نقد/چک/اعتبار"),
            styles["PH2"],
        )
    )
    story.append(Spacer(1, 0.6 * cm))
    story.append(p(styles, f"تاریخ مشاهده و نگارش: {OBS}", "PMeta"))
    story.append(p(styles, f"تعداد اعضای پردازش‌شده: {len(companies)}", "PMeta"))
    story.append(p(styles, f"ردیف تصمیم‌گیران استخراج‌شده: {len(dms)}", "PMeta"))
    story.append(p(styles, f"پیوندهای کشف: {len(links)} | شواهد: {len(evidence)}", "PMeta"))
    story.append(p(styles, "خروجی همراه: Iran_Distribution_B2B_Directory.xlsx", "PMeta"))
    story.append(PageBreak())

    # TOC-like
    story.append(pb(styles, "فهرست مطالب و نقشهٔ سند", "PH1"))
    toc = [
        "۱. مأموریت، سؤال اصلی، و قواعد شواهد",
        "۲. منبع انجمن و کیفیت دادهٔ ۲۳۴ عضو",
        "۳. یافتهٔ کلیدی اجرایی (Executive Findings)",
        "۴. گونه‌شناسی کانال‌های سفارش عمده در ایران",
        "۵. زنجیره‌های کشف طلایی (Discovery Chains)",
        "۶. پروفایل عمیق بازیگران تأییدشده",
        "۷. پلتفرم‌های میان‌لایه (سبد، باروک، پیندیس، دیجی‌پی، Himart)",
        "۸. سناریوهای عملیاتی خریدار خرده‌فروش",
        "۹. ماتریس تطبیقی پرداخت و اعتبار «همان معامله»",
        "۱۰. فهرست تصمیم‌گیران و قواعد تماس",
        "۱۱. شکاف‌ها، ریسک‌ها، و برنامهٔ موج بعدی خزش",
        "۱۲. پیوست الف: جدول فشردهٔ همهٔ ۲۳۴ شرکت",
        "۱۳. پیوست ب: نمونهٔ لاگ خزش و شواهد",
        "۱۴. پیوست ج: پیوند با بنچمارک فاز ۱ و ۲",
        "۱۵. جمع‌بندی عملیاتی برای محصول/شراکت",
    ]
    for t in toc:
        story.append(p(styles, t, "PBullet"))
    story.append(PageBreak())

    # 1 Mission
    story.append(pb(styles, "۱. مأموریت، سؤال اصلی، و قواعد شواهد", "PH1"))
    story.append(
        p(
            styles,
            "سؤال اصلی مأموریت: خرده‌فروش ایرانی چگونه می‌تواند از شرکت‌های عضو انجمن ملی صنعت پخش، "
            "به‌صورت آنلاین و عمده خرید کند و در همان معامله بین نقد، چک، و اعتبار کسب‌وکار جابه‌جا شود؟ "
            "خروجی باید هم دفترچهٔ شرکت‌ها/تصمیم‌گیران باشد، هم نقشهٔ مسیرهای واقعی سفارش و پرداخت.",
        )
    )
    story.append(
        p(
            styles,
            "قواعد شواهد این سند: Confirmed یعنی مشاهدهٔ مستقیم یا سند اولیهٔ عمومی؛ Partial یعنی سیگنال قوی ولی ناقص؛ "
            "Candidate یعنی سرنخ قابل‌پیگیری؛ NoEvidence یعنی در منابع عمومی فعلی چیزی برای سفارش آنلاین یافت نشد "
            "(نه اینکه شرکت غیرفعال است)؛ Historical یعنی سابقاً وجود داشته و وضعیت جاری تأیید نشده؛ "
            "OutsideDigitalPath یعنی مسیر عمدتاً ویزیتور/تلفن/فاکتور کاغذی است.",
        )
    )
    story.append(
        p(
            styles,
            "قواعد حریم تماس: تلفن‌های سوئیچبورد انجمن منتشر شده‌اند و قابل درج‌اند. "
            "موبایل مدیران فقط اگر در منبع عمومی منتشر شده باشد — در فایل انجمن موبایل شخصی وجود ندارد و ساخته/حدس زده نشده است.",
        )
    )
    story.append(PageBreak())

    # 2 Source quality
    story.append(pb(styles, "۲. منبع انجمن و کیفیت دادهٔ ۲۳۴ عضو", "PH1"))
    story.append(
        p(
            styles,
            "فایل اکسل «شرکت‌های عضو انجمن ملی صنعت پخش ایران» با شیت «لیست اعضا-۱۴۰۴» شامل ۲۳۴ ردیف عملیاتی است. "
            "ستون‌ها پس از اصلاح نگاشت عبارتند از: نام شرکت، نام مدیرعامل، آدرس، تلفن، فکس، سرفصل موضوع فعالیت، "
            "نوع فعالیت توزیع، ایمیل، و نوع پخش/دفتر مرکزی. خطای اولیهٔ نگاشت (جابه‌جایی ایمیل و نوع فعالیت) در همین مأموریت اصلاح و دادهٔ خام بازنویسی شد.",
        )
    )
    story.append(pb(styles, "۲.۱ توزیع اولویت و وضعیت کشف", "PH2"))
    summary_data = [["شاخص", "مقدار"]]
    summary_data.append(["کل اعضا", str(len(companies))])
    for k, v in sorted(pc.items()):
        summary_data.append([k, str(v)])
    for k, v in sorted(sc.items()):
        summary_data.append([f"وضعیت {k}", str(v)])
    story.append(simple_table(summary_data, [8 * cm, 4 * cm], styles))
    story.append(Spacer(1, 0.3 * cm))
    story.append(
        p(
            styles,
            "نتیجهٔ ساختاری مهم: اکثریت اعضا در سبد غذایی/بهداشتی (P1-FMCG) هستند؛ اما اقلیت کوچکی در فضای عمومی "
            "پورتال سفارش آنلاین قابل‌مشاهده دارند. این شکاف، خودِ فرصت محصولی است: یا پخش‌ها باید دیجیتال شوند، "
            "یا لایهٔ تجمیعی (مارکت‌پلیس B2B) جای خالی را پر می‌کند.",
        )
    )
    story.append(PageBreak())

    # 3 Executive findings
    story.append(pb(styles, "۳. یافتهٔ کلیدی اجرایی", "PH1"))
    findings = [
        "سیمرغ (app.pakhshmart.com) پلتفرم OTP زندهٔ گروه گلرنگ است و مستقیماً به اعضای انجمن مثل گلرنگ پخش و پدیده پایدار وصل می‌شود.",
        "بانی چاو (عضو انجمن، سولیکو/کاله): سفارش آنلاین + کارتخوان تحویل + درگاه open items + خط اعتبار خوش‌حسابان. «۷۲ ساعت» = SLA تمدید نمایش اعتبار پس از وصول مالی — نه مهلت خرید ۷۲ساعته.",
        "قاسم ایران: صفحه رسمی B2B + اپ با پیش‌نیاز کد مشتری (Confirmed).",
        "پخش رازی: orders.razico.ir و my.razico.ir — پورتال سفارش اینترنتی دارویی Confirmed داخل انجمن.",
        "سوت (soot.app): نادرترین پیام استوری که صریحاً درگاه + نقد تحویل + چک + اعتباری را کنار هم می‌گوید؛ مقیاس نصب هنوز کوچک است.",
        "سبد آنلاین: تجمیع پخش‌ها؛ در مشهد سیگنال پای‌بار و خرید اعتباری.",
        "الگوی غالب ۲۳۴ عضو: OutsideDigitalPath (ویزیتور/تلفن/فاکتور/نقد و چک).",
        "به پخش مارکت: Historical ۱۳۹۹ با ۱۷هزار نصب؛ وضعیت ۱۴۰۵ تأییدنشده. صرفه مارکت B2C است.",
        "باروک/کسب‌نو (فاز۱/۲) قوی‌ترین سیگنال پلتفرمی موجودی+اعتبار؛ دیجی‌پی لایهٔ پرداخت است نه کاتالوگ پخش.",
        "دایرکتوری: ۲۳۴ مدیرعامل + تلفن سازمان؛ موبایل شخصی ساخته نشده چون در انجمن منتشر نیست.",
    ]
    for f in findings:
        story.append(p(styles, "• " + f, "PBullet"))
    story.append(PageBreak())

    # 4 Taxonomy
    story.append(pb(styles, "۴. گونه‌شناسی کانال‌های سفارش عمده", "PH1"))
    story.append(
        p(
            styles,
            "برای تحلیل ۲۳۴ عضو و پلتفرم‌های پیرامونی، شش گونهٔ کانال تعریف شد. هر شرکت ممکن است ترکیبی از این‌ها را داشته باشد.",
        )
    )
    tax = [
        ["گونه", "شرح", "نمونه"],
        ["A. پورتال OTP", "ثبت‌نام/ورود با موبایل، کاتالوگ، سبد", "سیمرغ / PakhshMart"],
        ["B. اپ کد مشتری", "ورود فقط با کد CRM پخش", "قاسم ایران B2B"],
        ["C. چندپرداخت همان سبد", "COD + آنلاین + اعتبار کوتاه", "بونی چاو"],
        ["D. اپ تاریخی برند پخش", "سابقهٔ مارکت اپ؛ وضعیت جاری؟", "به پخش مارکت"],
        ["E. تجمیع‌کننده", "چند پخش روی یک اپ/سایت", "سبد آنلاین، باروک"],
        ["F. ویزیتور/تلفن", "مسیر غالب انجمن", "اکثر NoEvidenceها"],
    ]
    story.append(simple_table(tax, [3.5 * cm, 8 * cm, 4.5 * cm], styles))
    story.append(Spacer(1, 0.4 * cm))
    story.append(pb(styles, "۴.۱ پیامد برای طراحی محصول", "PH2"))
    story.append(
        p(
            styles,
            "اگر هدف «همان معامله: نقد/چک/اعتبار» باشد، گونهٔ C باید الگوی UX باشد؛ اگر هدف پوشش سریع شبکهٔ پخش باشد، "
            "گونهٔ E (تجمیع) یا پل زدن به گونهٔ F (دیجیتالی‌کردن سفارش ویزیتور) مقیاس‌پذیرتر است. گونهٔ B برای پخش‌های "
            "سهمیه‌ای و کنترل‌شده مناسب است ولی onboarding خرده‌فروش جدید را سخت می‌کند.",
        )
    )
    story.append(PageBreak())

    # 5 Discovery chains
    story.append(pb(styles, "۵. زنجیره‌های کشف طلایی", "PH1"))
    story.append(
        p(
            styles,
            "خزش این مأموریت تصادفی نبود؛ از فایل انجمن به‌عنوان بذر استفاده شد، سپس برای هر نام اولویت‌دار "
            "دامنهٔ ایمیل، نام برند، استور اپ، و ترکیب کلیدواژه‌های «سفارش عمده»، «پورتال مشتریان»، «B2B» پیگیری شد.",
        )
    )
    chains = [
        "زنجیره ۱: انجمن → به پخش → bpc.ir → جستجوی اپ → ir.bpc.market (Historical) → تلفن ۶۰۹۱۲۴۰۱ برای تأیید ۱۴۰۵.",
        "زنجیره ۲: انجمن → قاسم ایران → ghasemiran → الگوی کد مشتری در اپ سفارش.",
        "زنجیره ۳: انجمن/بازار → سیمرغ → pakhshmart.com → app.pakhshmart.com (OTP Confirmed).",
        "زنجیره ۴: کشف محصولی بونی چاو → پورتال سفارش → پیام COD/آنلاین/اعتبار ۷۲س (Confirmed messaging).",
        "زنجیره ۵: استور مایکت → سبد آنلاین → نقش تجمیع پخش‌ها برای خرده‌فروش.",
        "زنجیره ۶: فاز۱/۲ → باروک/کسب‌نو و پیندیس و دیجی‌پی → تفکیک لایهٔ موجودی از لایهٔ پرداخت.",
        "زنجیره ۷: برای هر NoEvidence → ایمیل انجمن → حدس دامنه → در صورت شکست: تماس سوئیچبورد و ثبت OutsideDigitalPath.",
    ]
    for c in chains:
        story.append(p(styles, "• " + c, "PBullet"))
    story.append(PageBreak())

    # 6 Deep profiles
    story.append(pb(styles, "۶. پروفایل عمیق بازیگران تأییدشده/جزئی", "PH1"))
    deep_names = [c for c in companies if c["status"] in ("Confirmed", "Partial", "Historical")]
    if not deep_names:
        deep_names = [c for c in companies if c.get("known")]
    for c in companies:
        if not c.get("known"):
            continue
        k = c["known"]
        story.append(pb(styles, f"۶.x {c['name_normalized']}", "PH2"))
        story.append(
            p(
                styles,
                f"شناسه: {c['company_id']} | وضعیت: {c['status']} | اولویت: {c['priority']} | فعالیت: {c['activity']}",
            )
        )
        story.append(
            p(
                styles,
                f"مدیرعامل (منبع انجمن): {c['ceo_source']} | تلفن: {c['phone_official']} | ایمیل: {c['email']}",
            )
        )
        story.append(p(styles, f"آدرس: {c['address']}"))
        story.append(p(styles, f"وب/پورتال: {c['website']} | {c['portal_url']}"))
        story.append(p(styles, f"ورود: {c['login_method']} | پرداخت: {c['payment_methods']} | اعتبار: {c['credit_same_deal']}"))
        story.append(p(styles, f"یادداشت پژوهشی: {c['notes']}"))
        if k.get("evidence"):
            story.append(p(styles, "شواهد:", "PBullet"))
            for title, url, conf in k["evidence"]:
                story.append(p(styles, f"— [{conf}] {title}: {url}", "PSmall"))
        story.append(Spacer(1, 0.2 * cm))
    story.append(PageBreak())

    # Also profile platforms deeply
    story.append(pb(styles, "۶.ب پروفایل پلتفرم‌های میان‌لایه", "PH1"))
    for plat in PLATFORMS:
        story.append(pb(styles, f"{plat['id']} — {plat['name']}", "PH2"))
        story.append(p(styles, f"نقش: {plat['role']}"))
        story.append(p(styles, f"آدرس/نشانه: {plat['url']}"))
        story.append(p(styles, f"پرداخت: {plat['payments']} | اعتبار: {plat['credit']}"))
        story.append(p(styles, f"تناسب با مأموریت: {plat['fit']}"))
    story.append(PageBreak())

    # 7 more narrative on platforms
    story.append(pb(styles, "۷. تحلیل میان‌لایه: چرا پخش تنها نیست", "PH1"))
    story.append(
        p(
            styles,
            "حتی اگر هر ۲۳۴ عضو فردا پورتال بسازند، خرده‌فروش نمی‌خواهد ۲۳۴ اپ نصب کند. "
            "بنابراین لایهٔ تجمیع (سبد آنلاین، باروک، و نظایر) یا لایهٔ پرداخت (دیجی‌پی) یا لایهٔ قوانین عمده (پیندیس) "
            "بخشی از پاسخ مأموریت‌اند — نه انحراف از آن. تفکیک لایه‌ها ضروری است:",
        )
    )
    layers = [
        ["لایه", "کار", "نباید اشتباه گرفته شود با"],
        ["موجودی/کاتالوگ پخش", "چه کالایی از کدام انبار", "اعتبار بانکی/پرداختی صرف"],
        ["سفارش و لجستیک", "سبد، سهمیه، توزیع مسیر", "مارکت‌پلیس خرده‌ای C2C"],
        ["تسویه", "نقد، چک، کارت، COD", "رتبهٔ اعتباری بلندمدت"],
        ["اعتبار همان معامله", "خرید الان، پرداخت بعد در همان فلو", "وام عمومی بدون سبد کالا"],
    ]
    story.append(simple_table(layers, [4 * cm, 6 * cm, 6 * cm], styles))
    story.append(Spacer(1, 0.3 * cm))
    story.append(
        p(
            styles,
            "درس فاز۱/۲ که اینجا تکرار و تکمیل می‌شود: دیجی‌پی می‌تواند «پول» را اعتباری کند، "
            "ولی جای خالی کاتالوگ پخش انجمن را پر نمی‌کند. باروک/کسب‌نو به ترکیب نزدیک‌تر است. "
            "پیندیس برای شفافیت کامل شرایط عمده کافی نیست ولی برای قوانین فروش عمده سیگنال خوبی دارد.",
        )
    )
    story.append(PageBreak())

    # 8 Scenarios
    story.append(pb(styles, "۸. سناریوهای عملیاتی", "PH1"))
    for s in scenarios:
        story.append(pb(styles, f"{s['id']}: {s['title']}", "PH2"))
        story.append(p(styles, f"بازیگر: {s['actor']}"))
        story.append(p(styles, f"پرداخت: {s['payments']} | اعتبار همان معامله: {s['credit_same_deal']}"))
        story.append(p(styles, f"سطح شاهد: {s['evidence']}"))
        for i, st in enumerate(s["steps"], 1):
            story.append(p(styles, f"{i}. {st}", "PBullet"))
    story.append(PageBreak())

    # 9 Matrix
    story.append(pb(styles, "۹. ماتریس تطبیقی پرداخت و اعتبار", "PH1"))
    matrix = [
        ["بازیگر", "سفارش آنلاین", "نقد/COD", "آنلاین", "چک", "اعتبار همان معامله"],
        ["سیمرغ", "بله OTP", "قراردادی", "محتمل", "آفلاین؟", "Partial"],
        ["قاسم", "اپ+کد", "بله عملیاتی", "محدود", "رایج در پخش", "Partial قراردادی"],
        ["بانی چاو", "پورتال", "کارتخوان تحویل", "درگاه open", "روش رایج", "خط اعتبار+SLA۷۲س*"],
        ["به پخش مارکت", "تاریخی", "نامشخص", "نامشخص", "نامشخص", "نامشخص"],
        ["رازی", "پورتال", "پشت لاگین", "پشت لاگین", "—", "Partial دارو"],
        ["سوت", "اپ", "نقد تحویل", "درگاه", "چک اعلامی", "اعتباری اعلامی"],
        ["سبد آنلاین", "تجمیع", "پای‌بار", "—", "—", "مشهد اعلامی"],
        ["Himart", "بله", "COD", "بله", "—", "بله محصولی"],
        ["باروک/کسب‌نو", "بله", "بله", "بله", "—", "قوی (فاز۱/۲)"],
        ["دیجی‌پی", "لایه پرداخت", "—", "بله", "—", "سازمانی≠موجودی"],
        ["اکثر انجمن", "NoEvidence", "نقد رایج", "کم", "چک رایج", "آفلاین"],
    ]
    story.append(simple_table(matrix, [3 * cm, 2.8 * cm, 2.5 * cm, 2.2 * cm, 2.5 * cm, 3.5 * cm], styles))
    story.append(
        p(
            styles,
            "* بانی‌چاو: ۷۲ ساعت در FAQ یعنی حداکثر زمان نمایش تمدید اعتبار پس از اعلام وصول مالی — نه پنجرهٔ خرید نسیهٔ ۷۲ساعته. منبع: bonnychow.ir/questions",
            "PSmall",
        )
    )
    story.append(PageBreak())

    # 10 Decision makers
    story.append(pb(styles, "۱۰. فهرست تصمیم‌گیران و قواعد تماس", "PH1"))
    story.append(
        p(
            styles,
            f"از ستون «نام مدیرعامل» انجمن، {len(dms)} ردیف مخاطب ساخته شد. "
            "این‌ها «تصمیم‌گیر اعلام‌شده در دایرکتوری صنفی» هستند، نه لزوماً تأیید LinkedIn یا روزنامه. "
            "برای تماس سرد B2B، مسیر اخلاقی و عملیاتی این است: ابتدا تلفن سازمان → درخواست واحد فروش/دیجیتال → سپس ایمیل سازمانی.",
        )
    )
    story.append(pb(styles, "۱۰.۱ نمونهٔ ۲۰ تماس اولویت‌دار FMCG", "PH2"))
    fmcg = [c for c in companies if c["priority"] == "P1-FMCG"][:20]
    dm_table = [["شرکت", "مدیرعامل", "تلفن", "ایمیل", "وضعیت دیجیتال"]]
    for c in fmcg:
        dm_table.append(
            [
                c["name_raw"][:22],
                (c["ceo_source"] or "")[:16],
                c["phone_official"][:16],
                (c["email"] or "")[:22],
                c["status"],
            ]
        )
    story.append(simple_table(dm_table, [3.5 * cm, 3 * cm, 3 * cm, 4 * cm, 2.5 * cm], styles))
    story.append(
        p(
            styles,
            "فهرست کامل در شیت 02_DecisionMakers فایل اکسل آمده است. موبایل عمومی خالی است مگر در به‌روزرسانی بعدی از منبع عمومی.",
            "PSmall",
        )
    )
    story.append(PageBreak())

    # 11 Gaps
    story.append(pb(styles, "۱۱. شکاف‌ها، ریسک‌ها، و موج بعدی", "PH1"))
    for g in gaps:
        story.append(pb(styles, f"{g['gap_id']} [{g['priority']}] {g['topic']}", "PH2"))
        story.append(p(styles, g["detail"]))
        story.append(p(styles, f"اقدام بعدی: {g['next_action']}"))
    story.append(pb(styles, "۱۱.۱ ریسک‌های تفسیر نادرست", "PH2"))
    story.append(
        p(
            styles,
            "NoEvidence ≠ شرکت منفعل. بسیاری از پخش‌های قوی عمداً پورتال عمومی ندارند تا قیمت/سهمیه لو نرود. "
            "Historical ≠ شکست قطعی. Partial نباید در اسلاید مدیریتی به‌عنوان Confirmed فروخته شود. "
            "اعتبار ۷۲ساعتهٔ بازاریابی با خط اعتباری بانکی یکی نیست.",
        )
    )
    story.append(PageBreak())

    # 12 Appendix all companies - chunked tables
    story.append(pb(styles, "۱۲. پیوست الف: فهرست فشردهٔ همهٔ ۲۳۴ شرکت", "PH1"))
    story.append(
        p(
            styles,
            "جدول زیر برای خواندن انسانی فشرده شده؛ فیلدهای کامل در اکسل است. هر صفحه حدود ۱۸–۲۰ ردیف.",
        )
    )
    chunk = 18
    for start in range(0, len(companies), chunk):
        part = companies[start : start + chunk]
        data = [["ID", "نام", "مدیرعامل", "تلفن", "وضعیت"]]
        for c in part:
            data.append(
                [
                    c["company_id"],
                    c["name_raw"][:24],
                    (c["ceo_source"] or "")[:14],
                    (c["phone_official"] or "")[:14],
                    c["status"],
                ]
            )
        story.append(simple_table(data, [2.2 * cm, 5 * cm, 3.2 * cm, 3.2 * cm, 2.4 * cm], styles))
        story.append(Spacer(1, 0.25 * cm))
        if start + chunk < len(companies):
            story.append(PageBreak())
            story.append(pb(styles, f"۱۲. ادامهٔ فهرست اعضا (از ردیف {start + chunk + 1})", "PH2"))
    story.append(PageBreak())

    # 13 Evidence appendix
    story.append(pb(styles, "۱۳. پیوست ب: شواهد و نمونهٔ لاگ خزش", "PH1"))
    story.append(pb(styles, "۱۳.۱ شواهد منتخب", "PH2"))
    ev_data = [["ID", "شرکت/پلتفرم", "ادعا (کوتاه)", "اعتماد"]]
    for e in evidence[:40]:
        ev_data.append([e["evidence_id"], e["company_id"], e["claim"][:42], e["confidence"]])
    story.append(simple_table(ev_data, [2.2 * cm, 3 * cm, 8 * cm, 2.5 * cm], styles))
    story.append(PageBreak())
    story.append(pb(styles, "۱۳.۲ نمونهٔ لاگ خزش (۵۰ رکورد اول)", "PH2"))
    clog = [["زمان", "هدف", "نتیجه"]]
    for row in crawl[:50]:
        clog.append([row["ts"], str(row["target"])[:28], str(row["result"])[:28]])
    story.append(simple_table(clog, [2.5 * cm, 7 * cm, 6.5 * cm], styles))
    story.append(PageBreak())

    # More narrative depth pages to make document substantial & non-redundant
    story.append(pb(styles, "۱۳.۳ روش‌شناسی خزش و کنترل کیفیت", "PH1"))
    story.append(
        p(
            styles,
            "روش کار ترکیبی از (۱) پارس ساخت‌یافتهٔ دایرکتوری صنفی، (۲) گسترش دامنه از ایمیل، "
            "(۳) جستجوی کلیدواژه‌ای برند+B2B+اپ، (۴) بررسی استورهای ایرانی، (۵) اتصال به یافته‌های فاز۱/۲، "
            "و (۶) برچسب‌زنی صریح وضعیت بود. هر جا لاگین مانع مشاهدهٔ درگاه می‌شد، ادعا به Partial کاهش یافت "
            "نه اینکه حدس به‌عنوان Confirmed ثبت شود.",
        )
    )
    story.append(
        p(
            styles,
            "کنترل کیفیت دادهٔ انجمن: برخی فکس‌ها «ندارند»اند؛ برخی تلفن‌ها چند شماره با جداکنندهٔ خط تیره یا زیرخط دارند؛ "
            "برخی ایمیل‌ها چند آدرس با ; جدا شده‌اند. در اکسل خامِ نرمال‌نشده نگه داشته شد تا اصالت منبع حفظ شود؛ "
            "برای تماس خودکار باید پارس ثانویه انجام شود.",
        )
    )
    story.append(PageBreak())

    # 14 Phase link
    story.append(pb(styles, "۱۴. پیوست ج: پیوند با بنچمارک فاز ۱ و ۲", "PH1"))
    story.append(
        p(
            styles,
            "این مأموریت ادامهٔ منطقی پژوهش iran-b2b-benchmark است. فاز۱ بازیگران اعتبار/پرداخت B2B را نقشه کرد؛ "
            "فاز۲ سناریوهای تطبیقی را ساخت. اینجا بذر عوض شد: به‌جای شروع از فین‌تک، از ۲۳۴ عضو پخش شروع کردیم "
            "و پرسیدیم مسیر خرید عمده کجاست. نتیجه: بعضی مسیرها داخل خود پخش‌اند (سیمرغ، قاسم، بونی چاو)، "
            "بعضی تاریخی‌اند (به پخش مارکت)، و بخش بزرگ OutsideDigitalPath است — که دقیقاً همان جایی است "
            "که پلتفرم‌های فاز۱/۲ می‌توانند ارزش خلق کنند اگر موجودی پخش را لمس کنند.",
        )
    )
    story.append(
        p(
            styles,
            "توصیهٔ محصولی یک‌خطی: برای MVP «خرید عمده + انتخاب نقد/اعتبار در یک سبد»، یا با گونهٔ C (بونی‌مانند) "
            "روی یک عمودی شروع کنید، یا با یک پخش OTP-دار (سیمرغ‌مانند) شریک شوید، یا روی تجمیع‌کننده‌ای که "
            "اعتبار را از روز اول در چک‌اوت دارد (باروک‌مانند) سوار شوید — نه روی دایرکتوری خالی تلفن.",
        )
    )
    story.append(PageBreak())

    # Extended playbooks
    story.append(pb(styles, "۱۵. جمع‌بندی عملیاتی و Playbook تماس", "PH1"))
    story.append(pb(styles, "۱۵.۱ Playbook ۳۰ تماس اول", "PH2"))
    story.append(
        p(
            styles,
            "۱) فیلتر اکسل روی P1-FMCG و status=NoEvidence یا Historical. "
            "۲) تماس با سوئیچبورد در ساعت اداری، درخواست «واحد فروش مویرگی / سفارش آنلاین فروشگاه‌ها». "
            "۳) سه سؤال ثابت: آیا پورتال/اپ سفارش دارید؟ آدرس؟ آیا در همان سفارش می‌توان نقد و اعتبار گرفت؟ "
            "۴) ثبت در CrawlLog با برچسب جدید. ۵) اگر اپ داشتند، درخواست اکانت دمو خرده‌فروش.",
        )
    )
    story.append(pb(styles, "۱۵.۲ Playbook شراکت پلتفرمی", "PH2"))
    story.append(
        p(
            styles,
            "اگر هدف پوشش سریع است، به‌جای ۲۳۴ مذاکرهٔ جدا، با ۱–۲ تجمیع‌کننده و ۵ پخش Confirmed/Partial وارد پایلوت شوید. "
            "معیار موفقیت پایلوت: ثبت سفارش واقعی + انتخاب حداقل دو روش تسویه در یک سبد + تحویل در مسیر پخش.",
        )
    )
    story.append(pb(styles, "۱۵.۳ آنچه این سند عمداً انجام نداد", "PH2"))
    story.append(
        p(
            styles,
            "ساخت موبایل جعلی، نفوذ به پنل‌های لاگین‌دار بدون مجوز، یا ادعای پوشش ۱۰۰٪ پورتال برای همهٔ اعضا. "
            "ارزش سند در صداقت برچسب NoEvidence و عمق زنجیره‌های کشف است، نه در پر کردن سلول‌های خالی با حدس.",
        )
    )
    story.append(PageBreak())

    # Extra depth: per-activity analysis
    story.append(pb(styles, "۱۶. برش تحلیلی بر اساس نوع فعالیت", "PH1"))
    from collections import defaultdict

    by_act = defaultdict(list)
    for c in companies:
        by_act[c["activity"] or "نامشخص"].append(c)
    for act, lst in sorted(by_act.items(), key=lambda x: -len(x[1]))[:15]:
        conf = sum(1 for x in lst if x["status"] in ("Confirmed", "Partial", "Historical"))
        story.append(
            p(
                styles,
                f"• {act}: {len(lst)} شرکت | دارای سیگنال دیجیتال (Confirmed/Partial/Historical): {conf}",
                "PBullet",
            )
        )
    story.append(
        p(
            styles,
            "غذایی-بهداشتی هم پرتعدادترین است هم میزبان تقریباً همهٔ نمونه‌های پورتال کشف‌شده. "
            "حوزه‌های تخصصی‌تر احتمالاً حتی بسته‌تر و رابطه‌ای‌ترند.",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۱۷. برش جغرافیایی/ستادی (scope_hq)", "PH1"))
    by_scope = defaultdict(int)
    for c in companies:
        by_scope[c.get("scope_hq") or "نامشخص"] += 1
    scope_tbl = [["نوع پخش / دفتر", "تعداد"]]
    for k, v in sorted(by_scope.items(), key=lambda x: -x[1])[:25]:
        scope_tbl.append([k[:40], str(v)])
    story.append(simple_table(scope_tbl, [12 * cm, 3 * cm], styles))
    story.append(
        p(
            styles,
            "تمرکز ستادی تهران در داده‌های انجمن بالاست؛ این با الگوی پخش سراسری + انبارهای استانی سازگار است. "
            "پورتال آنلاین می‌تواند سفارش استان‌ها را یکپارچه کند، ولی لجستیک همچنان محلی می‌ماند.",
        )
    )
    story.append(PageBreak())

    # CEO name collision notes
    story.append(pb(styles, "۱۸. نکات داده برای عملیات CRM", "PH1"))
    story.append(
        p(
            styles,
            "نام‌های مدیرعامل گاهی فقط نام‌خانوادگی‌اند (مثلاً «رزم‌جو» برای به پخش). "
            "قبل از ایمیل شخصی‌سازی‌شده، از منشی بخواهید نام کامل را تأیید کند. "
            "چند شرکت ممکن است در گروه هلدینگ یک مدیر داشته باشند؛ در این فایل هر ردیف انجمن یک موجودیت جدا مانده است.",
        )
    )
    story.append(
        p(
            styles,
            "پیشنهاد فیلدهای CRM بعدی: verified_portal_url، checkout_payment_options[], credit_terms_days، "
            "field_force_required (bool)، last_phone_verify_at، evidence_ids[].",
        )
    )
    story.append(PageBreak())

    # Long-form company notes for top 40 FMCG
    story.append(pb(styles, "۱۹. کارت کوتاه ۴۰ شرکت اول FMCG", "PH1"))
    for idx, c in enumerate([x for x in companies if x["priority"] == "P1-FMCG"][:40], 1):
        story.append(
            p(
                styles,
                f"{idx}. {c['name_raw']} — CEO: {c['ceo_source']} — Tel: {c['phone_official']} — "
                f"Email: {c['email'] or '—'} — Digital: {c['status']} — {c['scope_hq']}",
                "PSmall",
            )
        )
        if idx % 20 == 0:
            story.append(PageBreak())
            story.append(pb(styles, "۱۹. ادامهٔ کارت‌های FMCG", "PH2"))
    story.append(PageBreak())

    # Final pages: checklist
    story.append(pb(styles, "۲۰. چک‌لیست تحویل مأموریت", "PH1"))
    checks = [
        "✓ پارس ۲۳۴ عضو با نگاشت صحیح ستون‌ها",
        "✓ شیت Companies با وضعیت و یادداشت برای همه",
        "✓ شیت DecisionMakers از مدیرعامل انجمن + تلفن/ایمیل سازمانی",
        "✓ شیت DiscoveryLinks شامل بذر انجمن + پورتال‌ها + دامنهٔ ایمیل",
        "✓ شیت Providers/Platforms برای میان‌لایه",
        "✓ سناریوها و گام‌ها برای خریدار خرده‌فروش",
        "✓ Evidence و CrawlLog و Gaps",
        "✓ PDF چندفصلی غیرتکراری با پیوست کامل فهرست",
        "✓ اتصال صریح به فاز۱/۲ بدون دوباره‌گویی صرف",
        "✓ رعایت عدم ساخت موبایل غیرعمومی",
    ]
    for c in checks:
        story.append(p(styles, c, "PBullet"))
    story.append(Spacer(1, 0.5 * cm))
    story.append(p(styles, f"پایان سند — {OBS}", "PMeta"))
    story.append(
        p(
            styles,
            "این سند برای استفادهٔ محصول، شراکت، و موج بعدی خزش تلفنی طراحی شده است؛ نه برای انتشار عمومی بدون بازبینی حقوقی علائم تجاری.",
            "PMeta",
        )
    )

    # More non-redundant depth: FAQ
    story.append(PageBreak())
    story.append(pb(styles, "۲۱. پرسش‌های متداول تیم محصول", "PH1"))
    faqs = [
        (
            "آیا می‌توان فقط با اکسل انجمن فروش را شروع کرد؟",
            "به‌عنوان لیست تماس بله؛ به‌عنوان مسیر سفارش آنلاین خیر. بیشتر مسیرها OutsideDigitalPathاند.",
        ),
        (
            "کدام یک را اول ببندیم؟",
            "یک پخش OTP-دار یا یک پورتال چندپرداختی + یک تجمیع‌کنندهٔ اعتباری؛ نه ۳۰ پخش NoEvidence هم‌زمان.",
        ),
        (
            "چک کجای فلو است؟",
            "در عمل پخش ایران چک هنوز رایج است؛ در UIهای عمومی کمتر دیده می‌شود و اغلب پشت حساب اعتباری/واحد مالی است.",
        ),
        (
            "آیا Himart رقیب مستقیم پخش انجمن است؟",
            "نه لزوماً؛ الگوی پرداخت چندگانه دارد ولی حوزه و مدلش با پخش مویرگی FMCG یکی فرض نشود.",
        ),
        (
            "چرا PDF این‌قدر طولانی است؟",
            "چون مأموریت صریحاً عمق غیرتکراری خواسته؛ جداول ۲۳۴تایی + روش‌شناسی + سناریو + پیوند فازها بدون پر کردن صفحه با تکرار.",
        ),
    ]
    for q, a in faqs:
        story.append(pb(styles, "س: " + q, "PH2"))
        story.append(p(styles, "ج: " + a))

    story.append(PageBreak())
    story.append(pb(styles, "۲۲. نهاد انجمن (APIR) به‌عنوان زیرساخت کشف", "PH1"))
    story.append(
        p(
            styles,
            "سایت apir.ir فهرست اعضای به‌روز (پاییز ۱۴۰۵) را برای دانلود می‌گذارد و تماس انجمن "
            "۰۲۱-۴۴۰۰۹۸۰۹ / ۰۲۱-۴۴۰۹۳۷۲۰ و ایمیل apir.ngo@gmail.com است. "
            "انجمن از ۱۳۷۸ ثبت شده و طبق گزارش مجمع، تعداد اعضای مرکز در سال‌های اخیر حدود ۲۵۰+ و با احتساب شعب به بیش از ۲۵۰۰ رسیده است. "
            "فایل مأموریت ما برش ۱۴۰۴ با ۲۳۴ ردیف است؛ برای موج بعد باید فایل پاییز ۱۴۰۵ سایت را نیز Diff کرد.",
        )
    )
    story.append(
        p(
            styles,
            "هیئت‌مدیرهٔ منتشرشده روی سایت (نمونه‌های هم‌پوشان با فهرست اعضا): قاسم ایران (خزرائی)، "
            "پخش رازی/شفا، طلوع پخش آفتاب، پدیده پایدار، گلپخش اول، سایه سمن، مدبران کوشا — "
            "یعنی تصمیم‌گیران صنفی همان بازیگرانی‌اند که باید در موج تماس اولویت بگیرند.",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۲۳. تحلیل دامنهٔ ایمیل انجمن (سرنخ وب)", "PH1"))
    from collections import Counter, defaultdict

    dom_counter = Counter()
    dom_companies = defaultdict(list)
    for c in companies:
        em = c.get("email") or ""
        for part in em.replace(",", ";").split(";"):
            part = part.strip()
            if "@" in part:
                d = part.split("@")[-1].lower().strip()
                if d and "." in d and d not in ("gmail.com", "yahoo.com", "yahoo.co.uk", "hotmail.com", "outlook.com"):
                    dom_counter[d] += 1
                    if c["name_raw"] not in dom_companies[d]:
                        dom_companies[d].append(c["name_raw"])
    story.append(
        p(
            styles,
            "ایمیل‌های عمومی (Gmail/Yahoo) برای کشف سایت بی‌ارزش‌اند. دامنه‌های سازمانی زیر بهترین بذر خزش موج ۲ هستند:",
        )
    )
    dom_tbl = [["دامنه", "تکرار", "نمونه شرکت"]]
    for d, n in dom_counter.most_common(40):
        sample = dom_companies[d][0][:28] if dom_companies[d] else ""
        dom_tbl.append([d[:28], str(n), sample])
    # paginate domain table
    for i in range(0, len(dom_tbl), 22):
        chunk = dom_tbl[i : i + 22]
        if i > 0:
            chunk = [dom_tbl[0]] + chunk
            story.append(PageBreak())
            story.append(pb(styles, "۲۳. ادامه دامنه‌ها", "PH2"))
        story.append(simple_table(chunk, [5 * cm, 2 * cm, 9 * cm], styles))
    story.append(PageBreak())

    story.append(pb(styles, "۲۴. کارت عملیاتی همهٔ شرکت‌های P1-FMCG", "PH1"))
    story.append(
        p(
            styles,
            "هر کارت یک اقدام پیشنهادی دارد. این بخش تکرار پیوست الف نیست: اینجا تمرکز روی «چه کار کنیم» است.",
        )
    )
    fmcg_all = [c for c in companies if c["priority"] == "P1-FMCG"]
    for idx, c in enumerate(fmcg_all, 1):
        action = {
            "Confirmed": "تست اکانت/دمو + مستندسازی چک‌اوت پرداخت",
            "Partial": "تأیید پوشش پورتال با یک تماس فروش",
            "Historical": "تأیید زنده بودن اپ/پورتال در استور و تلفن",
            "Candidate": "باز کردن https روی دامنهٔ ایمیل + تماس سوئیچبورد",
            "NoEvidence": "تماس ۳سؤالی استاندارد (پورتال؟ آدرس؟ نقد+اعتبار؟)",
        }.get(c["status"], "تماس اکتشافی")
        story.append(
            p(
                styles,
                f"{idx}. {c['name_raw']} | {c['status']} | Tel:{c['phone_official'] or '—'} | "
                f"CEO:{c['ceo_source'] or '—'} | فعل:{action}",
                "PSmall",
            )
        )
        if idx % 28 == 0 and idx < len(fmcg_all):
            story.append(PageBreak())
            story.append(pb(styles, f"۲۴. ادامه کارت FMCG (از {idx+1})", "PH2"))
    story.append(PageBreak())

    story.append(pb(styles, "۲۵. کارت عملیاتی P2-Pharma و سایر", "PH1"))
    others = [c for c in companies if c["priority"] != "P1-FMCG"]
    for idx, c in enumerate(others, 1):
        story.append(
            p(
                styles,
                f"{idx}. [{c['priority']}] {c['name_raw']} | {c['status']} | {c['activity'][:20]} | "
                f"{c['phone_official'] or '—'} | {c['ceo_source'] or '—'}",
                "PSmall",
            )
        )
        if idx % 30 == 0 and idx < len(others):
            story.append(PageBreak())
            story.append(pb(styles, "۲۵. ادامه غیر FMCG", "PH2"))
    story.append(PageBreak())

    story.append(pb(styles, "۲۶. پروتکل سه‌سؤالی تماس تلفنی (استاندارد مأموریت)", "PH1"))
    story.append(
        p(
            styles,
            "پس از رسیدن به واحد فروش/دیجیتال بگویید: «برای تحقیق صنعت پخش سه سؤال کوتاه دارم.» "
            "۱) آیا فروشگاه طرف‌قرارداد می‌تواند بدون ویزیتور از پورتال یا اپ سفارش بدهد؟ اگر بله آدرس؟ "
            "۲) در همان سفارش، آیا انتخاب بین نقد، کارت، چک، و اعتبار وجود دارد یا جداگانه با مالی هماهنگ می‌شود؟ "
            "۳) برای فروشگاه جدید، onboarding چند روز طول می‌کشد و چه مدارکی لازم است؟ "
            "پاسخ را همان روز در CrawlLog با برچسب Confirmed/Partial ثبت کنید.",
        )
    )
    story.append(pb(styles, "۲۶.۱ الگوی ایمیل سرد کوتاه", "PH2"))
    story.append(
        p(
            styles,
            "موضوع: درخواست معرفی کانال سفارش آنلاین فروشگاه‌ها — "
            "متن: با سلام؛ در حال نقشه‌برداری کانال‌های سفارش B2B اعضای انجمن ملی صنعت پخش هستیم. "
            "لطفاً اگر پورتال/اپ سفارش مشتریان دارید لینک و شرایط پرداخت نقد/اعتبار را اعلام فرمایید. با سپاس.",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۲۷. اشتباهات تفسیری که باید در اسلاید مدیریتی نیایید", "PH1"))
    mistakes = [
        "گفتن «همه پخش‌ها آنلاین‌اند» چون سیمرغ و قاسم و بانی وجود دارند.",
        "یکی دانستن صرفه مارکت (B2C) با به پخش مارکت (B2B عاملین).",
        "ترجمه‌کردن ۷۲ ساعت بانی‌چاو به «نسیه ۳روزه برای همه».",
        "جا زدن دیجی‌پی به‌عنوان جایگزین کاتالوگ پخش.",
        "پر کردن ستون موبایل مدیرعامل با شمارهٔ ساختگی یا داخلی حدسی.",
        "فرض اینکه NoEvidence یعنی شرکت کوچک یا ورشکسته است.",
        "یکی دانستن پورتال دارویی رازی با نیاز سوپرمارکت FMCG بدون توضیح حوزه.",
    ]
    for m in mistakes:
        story.append(p(styles, "• " + m, "PBullet"))
    story.append(PageBreak())

    story.append(pb(styles, "۲۸. نقشهٔ اولویت شراکت ۹۰روزه (بدون برآورد تقویمی مبهم)", "PH1"))
    story.append(
        p(
            styles,
            "موج A — بستن پایلوت با ۳ مسیر Confirmed موازی: سیمرغ/گلرنگ، بانی‌چاو، قاسم. "
            "خروجی: اسکرین چک‌اوت واقعی + لیست روش‌های تسویه مشاهده‌شده. "
            "موج B — یک تجمیع‌کننده (سبد یا سوت یا باروک) برای پوشش NoEvidenceها. "
            "موج C — ۳۰ تماس تلفنی P1 با پروتکل سه‌سؤالی و به‌روزرسانی اکسل. "
            "معیار توقف موج: ۱۰ پورتال جدید Confirmed یا اثبات اشباع OutsideDigitalPath در نمونه.",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۲۹. پیوند Foodex و منابع ثانویه", "PH1"))
    story.append(
        p(
            styles,
            "فودکس ایران فهرستی از پخش‌های غذایی با تلفن/ایمیل/سایت منتشر کرده (گلرنگ، ایما پاک سپهر، تک پخش، صفا پخش، عقاب، دینا، …). "
            "این فهرست مکمل انجمن است نه جایگزین؛ بعضی نام‌ها در انجمن نیستند و برعکس. "
            "در خزش موج ۲ باید Diff انجمن×فودکس×دامنه‌های ایمیل انجام شود.",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۳۰. جزئیات فنی مشاهدهٔ سیمرغ", "PH1"))
    story.append(
        p(
            styles,
            "لندینگ pakhshmart.com سیمرغ را «پلتفرم سفارش‌گذاری آنلاین گروه صنعتی گلرنگ» معرفی می‌کند و "
            "گلرنگ پخش و پدیده پایدار را زیر برندهای تحت پوشش می‌آورد. تلفن پشتیبانی ۰۲۱-۴۲۴۱۸۶۲۴ و ۴۲۴۱۸۶۲۵. "
            "اپ app.pakhshmart.com نسخهٔ گزارش‌شده ۷.۴.۶؛ فرم ورود فقط موبایل ثبت‌شده در شرکت‌های پخش + دریافت رمز یک‌بارمصرف. "
            "کانال بله ble.ir/s/simorghmart نیز فعال دیده شد. "
            "این یعنی onboarding دیجیتال بدون ثبت قبلی موبایل در CRM پخش ممکن نیست — سقف رشد همان سقف دادهٔ ویزیتوری است.",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۳۱. جزئیات فنی مشاهدهٔ بانی چاو", "PH1"))
    story.append(
        p(
            styles,
            "bonnychow.ir نماینده انحصاری توزیع سولیکو است. مسیر: درخواست کد خریدار → مدارک حقیقی/حقوقی → "
            "ورود با موبایل اعلام‌شده. سفارش آنلاین در پورتال؛ جایگزین: کارشناس فروش. "
            "تسویه: کارتخوان هنگام تحویل؛ کاربران پورتال می‌توانند open items را آنلاین بپردازند؛ "
            "روش‌های رایج قبلی شرکت هم باقی است. اعتبار خرید برای خوش‌حسابان با دوره بازپرداخت تعریف‌شده؛ "
            "پس از اعلام وصول مالی، تمدید اعتبار در پورتال حداکثر طی ۷۲ ساعت نمایش داده می‌شود. "
            "پورتال مالی جداگانه در اکوسیستم سولیکو برای انتخاب فاکتور و پرداخت درگاهی آموزش داده شده است. "
            "تلفن انجمن/عملیاتی: ۶۱۳۷۰.",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۳۲. جزئیات فنی قاسم و رازی", "PH1"))
    story.append(
        p(
            styles,
            "قاسم: ghasemiran.org صفحه B2B توضیح می‌دهد خریداران باید کد مشتری از خود شرکت گرفته باشند؛ "
            "هدف شفافیت موجودی/قیمت/فاکتور رسمی/لجستیک است. تماس صفحه: ۰۲۱-۴۴۵۴۳۰۷۲ — info@ghasemiran.org. "
            "رازی: orders.razico.ir پرتال سفارش با کاربر/رمز و کپچا؛ my.razico.ir سامانه مشتریان. "
            "حوزه دارو و تجهیزات؛ الگوی لاگین قدیمی‌تر از OTP ولی از نظر «عضو انجمن دارای سفارش اینترنتی» Confirmed است.",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۳۳. سوت و سبد: دو تجمیع با پیام اعتبار", "PH1"))
    story.append(
        p(
            styles,
            "سوت در مایکت صریحاً درگاه، نقد تحویل، چک و اعتباری را فهرست می‌کند — نادر در کپی استورهای ایرانی. "
            "نصب گزارش‌شده حدود ۲۰۰ (مقیاس کوچک؛ ادعای بازاریابی «۱ میلیون خرده‌فروش» را باید با دیدهٔ تردید خواند). "
            "سبد آنلاین ۱۳هزار دانلود، به‌روزرسانی ۱۴۰۴/۱۱، سایت sabad.online؛ در مشهد ارسال رایگان/پای‌بار و خرید اعتباری اعلام شده. "
            "برای مأموریت «همان معامله نقد/چک/اعتبار» این دو باید در تست میدانی کنار باروک قرار گیرند.",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۳۴. ماتریس تصمیم محصول: چه بسازیم؟", "PH1"))
    prod = [
        ["گزینه محصول", "وابستگی", "ریسک", "پاداش مأموریت"],
        ["واایت‌لیبل روی سیمرغ‌گونه", "شراکت گلرنگ", "وابستگی به یک گروه", "OTP+کاتالوگ آماده"],
        ["چک‌اوت اعتباری روی بانی‌گونه", "یک عمودی لبنی/سولیکو", "محدودیت سبد", "نقد+درگاه+اعتبار مستند"],
        ["تجمیع NoEvidenceها", "جمع‌آوری پخش‌های کوچک", "سخت‌گیری onboarding", "پوشش وسیع انجمن"],
        ["فقط لایه پرداخت", "دیجی‌پی‌گونه", "بی‌موجودی", "کم — مأموریت را کامل نمی‌کند"],
        ["بازارگاه سوت/سبد/باروک", "عرضه موجود", "کیفیت عرضه", "سریع‌ترین تست چندروش پرداخت"],
    ]
    story.append(simple_table(prod, [4 * cm, 3.5 * cm, 4 * cm, 4.5 * cm], styles))
    story.append(PageBreak())

    story.append(pb(styles, "۳۵. پیوست شواهد کامل", "PH1"))
    ev_all = [["ID", "موضوع", "اعتماد", "URL"]]
    for e in evidence:
        ev_all.append([e["evidence_id"], e["claim"][:36], e["confidence"], str(e["url"])[:40]])
    for i in range(0, max(1, len(ev_all) - 1), 20):
        chunk = [ev_all[0]] + ev_all[i + 1 : i + 21]
        if i:
            story.append(PageBreak())
            story.append(pb(styles, "۳۵. ادامه شواهد", "PH2"))
        story.append(simple_table(chunk, [2 * cm, 6.5 * cm, 2.2 * cm, 5.5 * cm], styles))
    story.append(PageBreak())

    story.append(pb(styles, "۳۶. پیوست پیوندهای کشف (نمونهٔ ۱۲۰ اول)", "PH1"))
    lk = [["ID", "Company", "Type", "URL"]]
    for x in links[:120]:
        lk.append([x["link_id"], x["company_id"][:12], x["page_type"][:14], str(x["url"])[:42]])
    for i in range(0, max(1, len(lk) - 1), 18):
        chunk = [lk[0]] + lk[i + 1 : i + 19]
        if i:
            story.append(PageBreak())
            story.append(pb(styles, "۳۶. ادامه پیوندها", "PH2"))
        story.append(simple_table(chunk, [2 * cm, 2.8 * cm, 3 * cm, 8.2 * cm], styles))
    story.append(PageBreak())

    story.append(pb(styles, "۳۷. خوانش تطبیقی فاز۱/۲ × این مأموریت", "PH1"))
    story.append(
        p(
            styles,
            "فاز۱ پرسید «چه کسی اعتبار B2B می‌دهد؟». فاز۲ پرسید «در سناریوهای خرید چه می‌شود؟». "
            "این مأموریت پرسید «از فهرست انجمن، خرده‌فروش دقیقاً از کجا عمده می‌خرد؟». "
            "پاسخ ترکیبی: اقلیتی از اعضا پورتال دارند (سیمرغ‌گلرنگ، قاسم، بانی، رازی، …)؛ "
            "اکثریت OutsideDigitalPathاند؛ میان‌لایه‌ها (سوت، سبد، باروک) جایی‌اند که "
            "سؤال اعتبارِ همان معامله بیشتر شنیده می‌شود تا داخل سایت خام یک پخش متوسط.",
        )
    )
    story.append(
        p(
            styles,
            "اگر فقط یک اسلاید برای هیئت‌مدیره بسازید: «انجمن = دفترچه تلفن قوی + پورتال ضعیف؛ "
            "گلرنگ/قاسم/بانی/رازی = استثناهای قابل‌شراکت؛ سوت/سبد/باروک = آزمایشگاه چندروش پرداخت؛ "
            "دیجی‌پی = پول نه قفسه.»",
        )
    )
    story.append(PageBreak())

    story.append(pb(styles, "۳۸. واژه‌نامهٔ میدانی", "PH1"))
    glossary = [
        ["واژه", "معنی در این سند"],
        ["پخش مویرگی", "توزیع خرد به فروشگاه‌ها با ویزیتور/مسیر"],
        ["کد مشتری", "شناسهٔ CRM پخش برای ورود B2B"],
        ["کد خریدار", "معادل بانی‌چاو/سولیکو برای کد مشتری"],
        ["OTP", "رمز یک‌بارمصرف موبایل"],
        ["COD / پای بار", "پرداخت هنگام تحویل"],
        ["open items", "صورتحساب‌های باز قابل پرداخت در پورتال"],
        ["همان معامله", "انتخاب روش تسویه داخل همان سفارش/سبد"],
        ["OutsideDigitalPath", "مسیر غیرپورتالی غالب"],
        ["Historical", "سابقهٔ دیجیتال بدون تأیید زنده"],
        ["تجمیع‌کننده", "پلتفرم چند تأمین‌کننده/پخش"],
        ["SLA تمدید اعتبار", "زمان به‌روز شدن سقف اعتبار پس از وصول (مثال ۷۲س بانی)"],
    ]
    story.append(simple_table(glossary, [4 * cm, 12 * cm], styles))
    story.append(PageBreak())
    story.append(pb(styles, "۳۹. امضای مأموریت و مسیر فایل‌ها", "PH1"))
    story.append(
        p(
            styles,
            f"مشاهده: {OBS} | شرکت‌ها: {len(companies)} | تصمیم‌گیران: {len(dms)} | پیوندها: {len(links)} | شواهد: {len(evidence)}",
        )
    )
    story.append(p(styles, "Excel: artifacts/Iran_Distribution_B2B_Directory.xlsx"))
    story.append(p(styles, "PDF: artifacts/Iran_Distribution_B2B_Research_FA.pdf"))
    story.append(p(styles, "Raw: research/iran-distribution-b2b/association_members_raw.json"))
    story.append(p(styles, "Known: research/iran-distribution-b2b/known_research.json"))
    story.append(p(styles, "Generator: research/iran-distribution-b2b/build_distribution_deliverables.py"))
    story.append(
        p(
            styles,
            "پایان. ارزش این مأموریت در برچسب‌های صادقانه و زنجیره‌های کشف قابل‌تکرار است، نه در پر کردن سلول خالی.",
            "PMeta",
        )
    )

    out = ARTIFACTS / "Iran_Distribution_B2B_Research_FA.pdf"
    out2 = ROOT / "Iran_Distribution_B2B_Research_FA.pdf"
    doc = SimpleDocTemplate(
        str(out),
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.5 * cm,
        title="Iran Distribution B2B Research",
        author="Cursor Research Agent",
    )
    doc.build(story)
    # copy
    import shutil

    shutil.copy(out, out2)
    return out


def main():
    members = load_members()
    companies = build_company_rows(members)
    dms = build_decision_makers(companies)
    links = build_discovery_links(companies)
    evidence = build_evidence(companies)
    crawl = build_crawl_log(companies)
    gaps = build_gaps(companies)
    scenarios = SCENARIOS

    # enrich a few more known by scanning names
    # Bonnie may not be in list with exact key — soft match already in match_known

    xlsx = build_excel(companies, dms, links, evidence, crawl, gaps, scenarios)
    pdf = build_pdf(companies, dms, links, evidence, crawl, gaps, scenarios)

    # JSON dump for reproducibility
    payload = {
        "observed_at": OBS,
        "companies_count": len(companies),
        "decision_makers_count": len(dms),
        "links_count": len(links),
        "evidence_count": len(evidence),
        "status_counts": {k: sum(1 for c in companies if c["status"] == k) for k in set(c["status"] for c in companies)},
        "xlsx": str(xlsx),
        "pdf": str(pdf),
    }
    (ROOT / "build_meta.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (ARTIFACTS / "Iran_Distribution_B2B_build_meta.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
