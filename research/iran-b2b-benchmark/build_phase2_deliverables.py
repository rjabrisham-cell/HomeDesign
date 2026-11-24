# -*- coding: utf-8 -*-
"""Phase-2 deliverables: XLSX scenario bank + comprehensive FA RTL PDF."""
from __future__ import annotations

import json
from pathlib import Path

import arabic_reshaper
from bidi.algorithm import get_display
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
    HRFlowable,
    KeepTogether,
)

ROOT = Path(__file__).resolve().parent
DATA = json.loads((ROOT / "scenarios_phase2.json").read_text(encoding="utf-8"))
ART = Path(r"D:\HomeDesign\artifacts")
ART.mkdir(parents=True, exist_ok=True)

FONT_DIR = Path(r"C:\Windows\Fonts")
pdfmetrics.registerFont(TTFont("Vazir", str(FONT_DIR / "Vazirmatn-Regular.ttf")))
pdfmetrics.registerFont(TTFont("Vazir-Bold", str(FONT_DIR / "Vazirmatn-Bold.ttf")))
pdfmetrics.registerFont(TTFont("Vazir-Semi", str(FONT_DIR / "Vazirmatn-SemiBold.ttf")))

OBS = DATA["meta"]["observation_date_jalali"]
GREG = DATA["meta"]["observation_date_gregorian"]


def fa(text: str) -> str:
    return get_display(arabic_reshaper.reshape(str(text)))


def link(url: str, label: str | None = None) -> str:
    label = label or url
    return f'<link href="{url}" color="blue"><u>{fa(label)}</u></link>'


# ---------- XLSX ----------
def build_xlsx():
    wb = Workbook()

    # Hypotheses
    ws = wb.active
    ws.title = "HypothesisTests"
    headers = ["ID", "Phase1 Claim", "Verdict", "Finding", "Evidence IDs"]
    ws.append(headers)
    for h in DATA["hypothesis_tests"]:
        ws.append([h["id"], h["claim_phase1"], h["verdict"], h["finding"], ", ".join(h["evidence_ids"])])

    # Scenarios summary
    ws2 = wb.create_sheet("Scenarios")
    ws2.append(
        [
            "Scenario ID",
            "Company",
            "Product",
            "Role",
            "Goal",
            "Channel",
            "Territory",
            "Step Count",
            "Outcome",
            "Money/Risk",
            "Confidence",
            "Access Boundary",
            "Exceptions",
            "Evidence (step labels)",
        ]
    )
    for s in DATA["scenarios"]:
        ev = "; ".join(f"{st['n']}:{st.get('evidence','')}" for st in s["steps"])
        ws2.append(
            [
                s["id"],
                s["company"],
                s["product"],
                s["role"],
                s["goal"],
                s["channel"],
                s["territory"],
                len(s["steps"]),
                s["outcome"],
                s["money"],
                s["confidence"],
                s["access_boundary"],
                " | ".join(s.get("exceptions") or []),
                ev,
            ]
        )

    # Steps detail
    ws3 = wb.create_sheet("ScenarioSteps")
    ws3.append(
        [
            "Scenario ID",
            "Step",
            "Pre-state",
            "User/System Action",
            "Fields",
            "System Response",
            "Decision Role",
            "Evidence ID",
            "Evidence Label",
        ]
    )
    for s in DATA["scenarios"]:
        for st in s["steps"]:
            ws3.append(
                [
                    s["id"],
                    st["n"],
                    st["pre"],
                    st["action"],
                    st["fields"],
                    st["system"],
                    st["decision"],
                    st.get("evidence", ""),
                    st.get("label", ""),
                ]
            )

    # Evidence registry
    ws4 = wb.create_sheet("Evidence")
    evidence = [
        ["E-PD-01", "https://pindis.com/", GREG, "پین‌دیس", "لندینگ آمار و ارزش پیشنهادی", "product_claim"],
        ["E-PD-02", "https://play.google.com/store/apps/details?id=com.pindis", GREG, "Google Play", "۱۸هزار خرده‌فروش؛ خرید اعتباری", "product_claim"],
        ["E-PINDIS-MYKET-01", "https://myket.ir/app/com.pindis", GREG, "مایکت", "فیلد ثبت‌نام؛ ۷هزار دانلود؛ نظر باروک ۶مهر۱۴۰۵", "product_claim+user_report"],
        ["E-PD-TERMS-01", "https://pindis.com/terms-and-conditions/", GREG, "قوانین پین‌دیس", "احراز، سبد≠رزرو، تحویل، کیفیت تأمین‌کننده، اعتبار غیرنقدی", "official_doc"],
        ["E-BM-01", "https://bonak.market/", GREG, "بنک مارکت", "کاتالوگ قیمت‌دار؛ انبار", "observed"],
        ["E-BB-01", "https://www.bonakbazar.com/", GREG, "بنک‌بازار", "حجره؛ درخواست؛ قیمت توافقی", "observed"],
        ["E-SS-01", "https://snapp.supply/", GREG, "اسنپ ساپلای", "HoReCa؛ آمار", "product_claim"],
        ["E-PK-01", "https://pakhshiar.ir/", GREG, "پخش‌یار", "بازارگاه؛ اعتباری/چکی ادعا؛ FAQ ناقص", "product_claim"],
        ["E-DK-01", "https://www.digikalabusiness.com/", GREG, "دیجی‌کالا بیزینس", "عمده سازمانی", "observed"],
        ["E-DP-01", "https://www.mydigipay.com/merchant-acquisition/", GREG, "دیجی‌پی", "پذیرندگی", "official_doc"],
        ["E-DP-02", "https://www.mydigipay.com/academy/faq/", GREG, "دیجی‌پی FAQ", "نکول بی‌اثر روی تسویه", "official_doc"],
        ["E-DP-ACADEMY-ONB-01", "https://www.mydigipay.com/academy/آموزش-ثبت-نام...", GREG, "آکادمی دیجی‌پی", "گام‌های ثبت‌نام و ابزار پذیرش", "official_doc"],
        ["E-DP-B2O-01", "https://www.mydigipay.com/landing/B2O/", GREG, "دیجی‌پی B2O", "رفاه+تدارکات سازمان", "official_doc"],
        ["E-DP-ORG-MAG-01", "https://www.mydigipay.com/mag/loans-and-credits/organization-credit/", GREG, "مجله دیجی‌پی", "اعتبار سازمانی کارکنان", "official_doc"],
        ["E-SP-01", "https://snapppay.ir/", GREG, "اسنپ‌پی", "BNPL مصرف‌کننده", "official_doc"],
        ["E-SP-02", "https://academy.snapppay.ir/2026/04/21/...", GREG, "آکادمی اسنپ‌پی", "۸ گام پذیرندگی", "official_doc"],
        ["E-SP-ACQ-01", "https://academy.snapppay.ir/merchant-acquisition/", GREG, "ثبت‌نام اسنپ‌پی", "۴گام+شاخه کانال", "official_doc"],
        ["E-TA-01", "https://taracredit.com/org-req/", GREG, "تارا", "سازمانی", "official_doc"],
        ["E-KP-01", "https://keepa.ir/", GREG, "کیپا", "کیف‌پول‌ها", "official_doc"],
        ["E-GP-01", "https://greenpay360.ir/", GREG, "گرین‌پی", "نوی‌پو پالیز", "official_doc"],
        ["E-NAVIPO-NEWS-01", "ilna/bankemruz/ana PR", GREG, "خبر/روابط‌عمومی بانک کشاورزی", "صدور ضمانت؛ بازارگاه؛ آمار", "secondary_news"],
        ["E-ECOTAM-01", "https://ecotam.ir/", GREG, "اکوتام بانک تجارت", "کسب‌نو FMCG با باروک", "official_doc"],
        ["E-BAROOK-01", "purson/sharghdaily bank trade echo", GREG, "توضیح طرح کسب‌نو", "مراحل ثبت تا کارت؛ سقف؛ گردان", "secondary_official_echo"],
        ["E-BAROOK-SEC-01", "https://sapegah.com/barook/", GREG, "منبع ثانویه باروک", "مدارک؛ چک ۱۲۵٪؛ کارمزد ۲٪", "secondary"],
        ["E-IMIS-01", "https://imis.ir/solutions/wholesale-online/", GREG, "آیمیس", "قیمت گروهی؛ پلکانی؛ چک", "product_claim"],
        ["E-NANO-01", "https://nano-soft.ir/marketer-ordering-application/", GREG, "نانو", "ویزیتور آفلاین", "product_claim"],
        ["E-KARA-01", "https://karaict.com/نرم-افزار-پخش-مویرگی/", GREG, "کارا", "سقف اعتبار مشتری", "product_claim"],
        ["E-BL-01", "https://baloan.ir/", GREG, "بالون", "BNPL", "official_doc"],
        ["E-AZ-01", "https://azkivam.com/", GREG, "ازکی‌وام", "وام کالا", "official_doc"],
        ["E-NEWS-SMP-01", "https://www.zoomit.ir/tech-iran/443896-snappfood-bulk-sale-launch-snappmarketpro/", "2025-07", "زومیت", "اسنپ‌مارکت‌پرو محدود تهران", "secondary_news"],
        ["SIM-01", "internal", GREG, "شبیه‌سازی پژوهش", "دوپخشی ۲۳۵٫۶م و مرجوعی", "simulation_design"],
    ]
    ws4.append(["Evidence ID", "URL/Source", "Date", "Source Name", "Shows", "Label"])
    for row in evidence:
        ws4.append(row)

    # Matrix
    ws5 = wb.create_sheet("ComparisonMatrix")
    ws5.append(
        [
            "Step",
            "Pindis",
            "Bonak Market",
            "Bonakbazar",
            "Pakhshiar",
            "Barook/Kasbno",
            "NaviPo",
            "Digipay Merchant",
            "Scenario IDs",
        ]
    )
    matrix_rows = [
        ["احراز خریدار کسب‌وکار", "بله قوانین/مایکت", "نامعلوم لاگین", "حجره/درخواست", "ادعا", "بله شعبه+وب", "بله پالیز+شعبه", "پذیرنده بله", "MP-PINDIS-REG-01;CR-BAROOK-ONB-01"],
        ["نمایش قیمت کارتنی", "ادعا+قوانین VAT", "مشاهده عمومی", "توافقی", "ادعا", "نزد تأمین‌کننده", "پالیزمارکت", "n/a", "MP-BONAK-ORDER-01;MP-PINDIS-ORDER-01"],
        ["مقایسه چند عرضه", "ادعا", "خیر", "بله آگهی", "ادعا", "شبکه تأمین", "تأمین‌کنندگان احراز", "n/a", "MP-PINDIS-ORDER-01;MP-BONAKBAZAR-RFQ-01"],
        ["سبد≠رزرو تا نهایی", "بله ماده۷", "نامعلوم", "n/a RFQ", "نامعلوم", "پس از پرداخت اعتبار", "پس از مصرف", "پرداخت لحظه‌ای", "MP-PINDIS-ORDER-01"],
        ["تخفیف پلکانی قاعده‌مند", "جشنواره/ادعا", "٪ نمایشی", "توافقی", "ادعا بدون متن", "نامعلوم", "نامعلوم", "کمپین مشتری", "SW-IMIS-01;MP-PAKHSHIAR-01"],
        ["پرداخت اعتباری همان سفارش موجودی", "اشاره+نظر باروک", "مشاهده نشد", "نامعلوم", "ادعا چکی", "بله هسته محصول", "بله", "خیر (مصرف‌کننده/سازمان)", "CR-BAROOK-BUY-01;CR-NAVIPO-01"],
        ["تسویه تضمینی فروشنده/پخش", "نامعلوم عمومی", "n/a خودفروش", "n/a", "نامعلوم", "بله ادعا بانک", "بله بانک", "بله FAQ مشتری‌نکول", "CR-BAROOK-BUY-01;CR-DIGIPAY-C-BNPL-01"],
        ["مرجوعی/کسری", "عدم حضور مستند؛ کسری ضعیف", "نامعلوم", "نامعلوم", "نامعلوم", "نامعلوم عمومی", "نامعلوم", "قراردادی", "MP-PINDIS-DELIVERY-01;REF-B2B-RETURN-01"],
        ["پنل ادمین داخلی", "ندیدیم", "ندیدیم", "ندیدیم", "ندیدیم", "ندیدیم", "پالیز ادعا", "ندیدیم؛ مرچنت دیدیم", "OPS-*; access boundaries"],
    ]
    for r in matrix_rows:
        ws5.append(r)

    # Best practices
    ws6 = wb.create_sheet("BestPractices")
    ws6.append(
        [
            "Practice",
            "Problem",
            "Rule",
            "Source Scenario",
            "Evidence",
            "Why selected",
            "Alternative",
            "Transfer condition to supermarket",
            "Ops load",
            "Metric",
            "Label",
        ]
    )
    practices = [
        [
            "اعتبار خرید موجودی با تضمین بانک برای تأمین‌کننده",
            "ترس پخش از چک برگشتی فروشگاه",
            "بانک نکول را می‌پذیرد؛ فروشگاه به بانک بدهکار است",
            "CR-BAROOK-BUY-01",
            "E-BAROOK-01;E-ECOTAM-01",
            "تنها الگوی نزدیک به سؤال محوری سوپر+موجودی",
            "نوی‌پو (صنف کشاورزی)",
            "نیاز شبکه تأمین‌کننده FMCG عضو طرح",
            "متوسط-بالا (شعبه+وثیقه)",
            "زمان تا اولین خرید اعتباری؛ نرخ نکول",
            "observed_pattern",
        ],
        [
            "سبد تا نهایی رزرو ایجاد نکند",
            "تعارض موجودی و قیمت لحظه‌ای",
            "افزودن به سبد حق ایجاد نمی‌کند؛ تغییرات تا ثبت نهایی اعمال می‌شود",
            "MP-PINDIS-ORDER-01",
            "E-PD-TERMS-01",
            "قاعده حقوقی صریح و قابل‌انتقال",
            "رزرو نرم با TTL",
            "همه بازارگاه‌ها",
            "کم",
            "نرخ لغو پس از نهایی؛ اختلاف موجودی",
            "official_doc",
        ],
        [
            "جداسازی تسویه پذیرنده از وصول اقساط مشتری",
            "مانع پذیرش فروش اعتباری",
            "نکول مشتری روی واریز فروشنده اثر ندارد",
            "CR-DIGIPAY-C-BNPL-01",
            "E-DP-02",
            "FAQ رسمی صریح",
            "اسنپ‌پی تسویه تضمینی ادعا",
            "برای BNPL مصرف‌کننده؛ برای پخش باید مدل بانک‌محور",
            "کم برای پذیرنده",
            "تأخیر تسویه؛ شکایت پذیرنده",
            "official_doc",
        ],
        [
            "مالکیت کیفیت و فاکتور تأمین‌کننده شفاف در قوانین",
            "ابهام اختلاف کالا",
            "کیفیت/پس از فروش با تأمین‌کننده؛ فاکتور همراه ارسال",
            "MP-PINDIS-DELIVERY-01",
            "E-PD-TERMS-01",
            "کاهش ابهام مسئولیت",
            "ضمانت پلتفرم مرکزی",
            "اگر چندعرضه‌کننده است ضروری",
            "متوسط پشتیبانی",
            "زمان رفع اختلاف کیفیت",
            "official_doc",
        ],
        [
            "نمایش کسری سقف اعتبار قبل از تأیید سفارش چندپخشی",
            "گیر کردن معامله ۲۳۵٫۶م روی سقف ۲۰۰م",
            "UI باید کسری ۳۵٫۶م و گزینه‌های نقد/کاهش/تقسیم را نشان دهد",
            "FIN-TWOSUPPLIER-SIM-01;REF-B2B-CORE-01",
            "SIM-01",
            "رفع اصطکاک مالی طراحی",
            "افزایش سقف خودکار",
            "هر محصول اعتبار B2B",
            "کم توسعه / زیاد محصول",
            "نرخ رها کردن سبد به‌خاطر سقف",
            "simulation_design",
        ],
        [
            "مرزبندی رفاه سازمانی از اعتبار موجودی فروشگاه",
            "اشتباه استراتژی محصول",
            "اگر بدهکار کارمند/سازمان رفاهی است، مدل را موجودی‌پخش ننام",
            "CR-DIGIPAY-ORG-01;CR-TARA-ORG-01",
            "E-DP-B2O-01;E-TA-01",
            "جلوگیری از نتیجه‌گیری غلط فاز۱",
            "تنخواه تدارکات فقط برای مصرف سازمان",
            "—",
            "—",
            "دقت طبقه‌بندی در تصمیم سرمایه‌گذاری",
            "analytical",
        ],
        [
            "مسیر مرجوعی قبل/بعد تسویه جدا",
            "مغایرت مالی پس از واریز به پخش",
            "قبل تسویه: کاهش فاکتور+آزادسازی اعتبار؛ بعد: سند اصلاحی سه‌طرفه",
            "REF-B2B-RETURN-01",
            "SIM + gap market",
            "بازار شاهد کامل ندارد؛ باید طراحی شود",
            "فقط پشتیبانی دستی",
            "ضروری برای مقیاس",
            "بالا در شروع",
            "زمان رفع مغایرت؛ نرخ برگشت پرونده",
            "simulation_design",
        ],
        [
            "ابزار ویزیتوری آفلاین مکمل بازارگاه",
            "پوشش پخش سنتی",
            "ثبت سر مغازه + همگام‌سازی",
            "SW-NANO-01;SW-KARA-01",
            "E-NANO-01;E-KARA-01",
            "واقعیت کانال پخش ایران",
            "فقط سلف‌سرویس اپ خریدار",
            "برای نفوذ سریع به پخش‌های موجود",
            "آموزش ویزیتور",
            "نسبت سفارش ویزیتوری/اپ",
            "product_claim",
        ],
    ]
    for p in practices:
        ws6.append(p)

    # Gaps
    ws7 = wb.create_sheet("Gaps")
    ws7.append(["Gap ID", "Topic", "What is missing", "Next research action"])
    gaps = [
        ["G1", "ادغام رسمی پین‌دیس×باروک", "صفحه/قرارداد مشترک مشاهده نشد؛ فقط نظر کاربر", "حساب آزمایشی یا مصاحبه تأمین‌کننده عضو هر دو"],
        ["G2", "قاعده پله تخفیف پخش‌یار", "FAQ بدون متن پاسخ در خزش", "دمو یا اسکرین قوانین قیمت"],
        ["G3", "کسری تحویل پین‌دیس", "فقط عدم حضور مستند است", "راهنمای مرجوعی داخل اپ"],
        ["G4", "کارمزد دقیق کسب‌نو/باروک", "منابع ثانویه ۲٪؛ ecotam «کمترین کارمزد»", "استعلام شعبه/قرارداد"],
        ["G5", "پنل ادمین داخلی همه بازیگران", "مشاهده نشد", "دمو رسمی نقش عملیات"],
        ["G6", "جیحون و آرین تدبیر", "محصول اعتباری احراز نشد", "بسته اسناد کارفرما"],
        ["G7", "اسنپ‌پی×ساپلای همان سفارش", "اتصال پرداخت احراز نشد", "تست حساب رستوران"],
        ["G8", "مرجع کیفیت بنک‌بازار", "مبهم در فرانت", "قوانین حجره/داوری"],
    ]
    for g in gaps:
        ws7.append(g)

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Border(
        left=Side(style="thin", color="B0B0B0"),
        right=Side(style="thin", color="B0B0B0"),
        top=Side(style="thin", color="B0B0B0"),
        bottom=Side(style="thin", color="B0B0B0"),
    )
    for sheet in wb.worksheets:
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = thin
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

    out = ROOT / "Iran-B2B-Scenario-Bank-Phase2.xlsx"
    wb.save(out)
    wb.save(ART / "Iran-B2B-Scenario-Bank-Phase2.xlsx")
    print("XLSX", out)
    return out


# ---------- PDF ----------
styles = getSampleStyleSheet()
for name, kwargs in {
    "FATitle": dict(fontName="Vazir-Bold", fontSize=16, leading=26, alignment=TA_CENTER, spaceAfter=10),
    "FAH1": dict(fontName="Vazir-Bold", fontSize=13, leading=22, alignment=TA_RIGHT, spaceBefore=12, spaceAfter=6, textColor=colors.HexColor("#15202b")),
    "FAH2": dict(fontName="Vazir-Semi", fontSize=11, leading=18, alignment=TA_RIGHT, spaceBefore=9, spaceAfter=4, textColor=colors.HexColor("#1f3347")),
    "FAH3": dict(fontName="Vazir-Semi", fontSize=10, leading=16, alignment=TA_RIGHT, spaceBefore=7, spaceAfter=3, textColor=colors.HexColor("#243b53")),
    "FABody": dict(fontName="Vazir", fontSize=9, leading=15, alignment=TA_JUSTIFY, spaceAfter=5),
    "FASmall": dict(fontName="Vazir", fontSize=7.8, leading=12, alignment=TA_RIGHT, spaceAfter=2),
    "FATag": dict(fontName="Vazir", fontSize=7.5, leading=11, alignment=TA_RIGHT, textColor=colors.HexColor("#6b3a1f"), spaceAfter=3),
    "FACaption": dict(fontName="Vazir", fontSize=7, leading=10, alignment=TA_CENTER, textColor=colors.HexColor("#555"), spaceAfter=6),
    "FACell": dict(fontName="Vazir", fontSize=7, leading=10, alignment=TA_RIGHT),
    "FACellBold": dict(fontName="Vazir-Bold", fontSize=7, leading=10, alignment=TA_RIGHT),
    "FACenter": dict(fontName="Vazir", fontSize=9, leading=15, alignment=TA_CENTER, spaceAfter=5),
}.items():
    styles.add(ParagraphStyle(name=name, **kwargs))


def P(t, style="FABody"):
    return Paragraph(fa(t), styles[style])


def H1(t):
    return Paragraph(fa(t), styles["FAH1"])


def H2(t):
    return Paragraph(fa(t), styles["FAH2"])


def H3(t):
    return Paragraph(fa(t), styles["FAH3"])


def hr():
    return HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#c5ced8"), spaceBefore=3, spaceAfter=6)


def tbl(headers, rows, widths):
    data = [[Paragraph(fa(h), styles["FACellBold"]) for h in headers]]
    for r in rows:
        data.append([Paragraph(fa(str(c)), styles["FACell"]) for c in r])
    t = Table(data, colWidths=widths, repeatRows=1)
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e7eef6")),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#aeb7c2")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f8fafc")]),
            ]
        )
    )
    return t


def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Vazir", 6.5)
    canvas.setFillColor(colors.HexColor("#666"))
    canvas.drawCentredString(A4[0] / 2, 1.15 * cm, fa(f"فاز ۲ — سناریوهای خرید عمده و اعتبار B2B ایران — صفحه {doc.page}"))
    canvas.drawString(1.5 * cm, A4[1] - 1.0 * cm, f"Obs {GREG}")
    canvas.drawRightString(A4[0] - 1.5 * cm, A4[1] - 1.0 * cm, fa(OBS))
    canvas.restoreState()


def build_pdf():
    story = []
    story.append(Spacer(1, 1.8 * cm))
    story.append(Paragraph(fa("بررسی رقبا و بهترین رویه‌های خرید عمده و اعتبار در ایران"), styles["FATitle"]))
    story.append(P("فاز ۲ — از فهرست رقبا تا سناریوی قابل‌اجرا", "FACenter"))
    story.append(P(f"تاریخ مشاهده: {OBS} ({GREG}) | بانک سناریو + ماتریس + مرجع طراحی", "FACenter"))
    story.append(hr())
    story.append(
        P(
            "سؤال محوری: وقتی فروشگاه مواد غذایی می‌خواهد موجودی را از پخش تأمین کند، مسیر واقعی از ورود تا تحویل و تسویه چیست؟ در کدام نمونه همین خرید عمده با اعتبار کسب‌وکاری انجام می‌شود و چگونه اعتبار و تسویه همان سفارش ردیابی می‌شود؟"
        )
    )
    story.append(PageBreak())

    # Exec summary
    story.append(H1("۱. خلاصهٔ مدیریتی — آزمون فرضیه‌های فاز ۱"))
    story.append(
        P(
            "فرضیهٔ جدایی کامل بازار سفارش و اعتبار اصلاح جزئی شد: نزدیک‌ترین معاملهٔ واحد مستند برای سوپرمارکت FMCG، طرح کسب‌نو بانک تجارت با تجمیع‌کننده باروک است (هویت کسب‌وکار + خرید از تأمین‌کننده عضو + مصرف اعتبار روی همان خرید + تسویه تضمینی بانک برای تأمین‌کننده). اتصال پرداخت باروک داخل پین‌دیس در نظر کاربر مایکت گزارش شده و صفحهٔ رسمی مشترک هنوز مشاهده نشده است. الگوی موازی قوی‌تر از نظر معماری یکپارچه، نوی‌پو/پالیز/گرین‌پی است ولی صنف کشاورزی/نهاده دارد."
        )
    )
    story.append(
        tbl(
            ["فرضیه", "حکم", "اثر تصمیم"],
            [
                ["H1 جدایی سفارش/اعتبار", "اصلاح جزئی", "برای موجودی سوپر سراغ باروک/کسب‌نو و نوی‌پو بروید نه فقط BNPL مصرف‌کننده"],
                ["H2 تفاوت مدل عملیاتی", "تأیید", "پرسش مالک قیمت/موجودی/فاکتور/پول قبل از برچسب بازارگاه"],
                ["H3 مرز رفاه/موجودی/پذیرندگی", "تیزتر شد", "دیجی‌پی سازمانی ≠ اعتبار خرید از پخش"],
                ["H4 آمار لندینگ", "تأیید", "۵۰هزار vs ۱۸هزار vs ۷هزار دانلود را با تعریف شاخص جدا گزارش کنید"],
            ],
            [3.2 * cm, 2.5 * cm, 10.8 * cm],
        )
    )
    story.append(Paragraph(fa("جدول الف — حکم فرضیه‌ها"), styles["FACaption"]))
    story.append(H2("تصمیم‌های قابل دفاع با شاهد فعلی"))
    for b in [
        "اگر هدف «خرید موجودی فروشگاه با اعتبار» است، الگوی مرجع بازار باروک/کسب‌نو است نه تارا/کیپا/BNPL خرد.",
        "اگر هدف «بازارگاه چندپخشی سلف‌سرویس» است، پین‌دیس نزدیک‌ترین شاهد قوانین+عملیات است؛ عنوان بازارگاه چندفروشنده فقط با مشاهدهٔ چند عرضهٔ مستقل در یک سبد قطعی می‌شود (پشت لاگین مانده).",
        "دیجی‌پی/اسنپ‌پی بهترین الگوی پذیرندگی و جداسازی ریسک نکول مشتری از تسویه فروشنده‌اند؛ برای پخش باید به مدل بانک‌محور ترجمه شوند.",
        "سناریوی مرجع طراحی (REF-B2B-CORE-01) مسیر پایه+اعتبار+مرجوعی را جمع می‌کند و از مشاهدهٔ خام متمایز برچسب خورده است.",
    ]:
        story.append(P("• " + b))
    story.append(PageBreak())

    # Method
    story.append(H1("۲. روش، دسترسی و نقشهٔ به‌روز"))
    story.append(
        P(
            f"پنجرهٔ مشاهده {GREG}. منابع: وب رسمی، قوانین، آکادمی پذیرندگان، فروشگاه اپ، صفحات بانک/اکوتام، خبر روابط‌عمومی، و نظر کاربر منتشرشده. حساب کامل خریدار/پخش/ادمین در اختیار نبود؛ مرز دسترسی در هر سناریو آمده است. هیچ دور زدن دسترسی انجام نشد."
        )
    )
    story.append(
        tbl(
            ["بازیگر", "چه کسی قیمت می‌دهد؟", "مالک موجودی؟", "تأیید معامله؟", "فاکتور؟", "گیرنده پول؟", "تحویل/مرجوعی؟"],
            [
                ["پین‌دیس", "سامانه/تأمین‌کننده", "ترکیبی (خرید پلتفرم یا عرضه تأمین‌کننده)", "ثبت نهایی خریدار+پردازش", "تأمین‌کننده همراه ارسال", "عامل پرداخت تعریف‌شده", "پلتفرم ارسال؛ کیفیت با تأمین‌کننده؛ عدم حضور→مرجوع کاربر"],
                ["بنک مارکت", "خود شرکت", "انبار خود", "خود شرکت", "خود شرکت", "خود شرکت", "از انبار ادعا"],
                ["بنک‌بازار", "توافقی حجره", "تأمین‌کننده آگهی", "طرفین", "نامعلوم عمومی", "احتمالاً مستقیم", "نامعلوم؛ مرجع کیفیت مبهم"],
                ["اسنپ ساپلای", "پلتفرم/تأمین", "تأمین‌کننده/عملیات", "خریدار+ساپلای", "نامعلوم", "ساپلای/تأمین", "تحویل زمان‌بندی"],
                ["پخش‌یار", "پخش‌ها ادعا", "پخش ادعا", "نامعلوم", "نامعلوم", "پخش/پلتفرم؟", "ارسال مستقیم ادعا"],
                ["دیجی‌کالا بیزینس", "دیجی‌کالا", "مدل دیجی‌کالا", "بیزینس", "دیجی‌کالا", "دیجی‌کالا", "لجستیک دیجی‌کالا"],
                ["آیمیس/نانو/کارا", "شرکت پخش مشتری نرم‌افزار", "پخش", "مدیر/ویزیتور", "پخش", "پخش", "پخش"],
                ["باروک کسب‌نو", "تأمین‌کننده عضو", "تأمین‌کننده", "خرید+پرداخت اعتبار", "تأمین‌کننده", "بانک/طرح به تأمین‌کننده", "تأمین‌کننده؛ ریسک نکول بانک ادعا"],
                ["نوی‌پو", "تأمین‌کننده احراز پالیز", "تأمین‌کننده", "بانک+سامانه", "زنجیره", "بانک/کیف", "حسب زنجیره"],
            ],
            [2.2 * cm, 2.3 * cm, 2.5 * cm, 2.3 * cm, 2.2 * cm, 2.5 * cm, 2.6 * cm],
        )
    )
    story.append(Paragraph(fa("جدول ب — مدل عملیاتی با پرسش واحد"), styles["FACaption"]))
    story.append(
        P(
            "اختلاف ارقام پین‌دیس: سایت «۵۰هزار+ خرده‌فروش» و «۶۵۰هزار+ سفارش» (ادعای شرکت، E-PD-01)؛ پلی‌استور/مایکت «۱۸هزار خرده‌فروش» (E-PD-02/E-PINDIS-MYKET-01)؛ مایکت «۷هزار دانلود». این‌ها تعریف‌های متفاوت‌اند و سهم بازار مستند نیستند."
        )
    )
    story.append(PageBreak())

    # International abstract
    story.append(H1("۲.۵ Executive Abstract (EN)"))
    story.append(
        P(
            "Research question: How does an Iranian grocery retailer actually procure inventory from distributors, and where can the same wholesale order be financed with business credit whose settlement is traceable? Method: public-web journey reconstruction (official sites, terms, merchant academies, bank SCF pages, app-store listings, dated news), labeled by evidence type; no unauthorized access. Result: consumer BNPL (Digipay/Snapp Pay) and employee-welfare credit (Tara/Keepa/Digipay B2O) are not substitutes for retailer inventory credit. The closest documented unified pattern for supermarket FMCG is Tejarat Bank’s Kasbno product via Barook aggregator (business KYC → revolving limit → purchase at contracted suppliers → bank-backed supplier settlement → 1–6 month repayment). A user report links Barook payment inside PinDis; a joint official page was not observed. Green Pay’s NaviPo/Paliz is a stronger integrated goods+credit architecture but for agri/livestock supply chains. PinDis terms provide the strongest public wholesale journey rules (cart≠reservation, supplier quality liability, no-show return). Deliverables: 32 distinct scenarios, editable XLSX, comparison matrix, reference design journey REF-B2B-CORE-01."
        )
    )
    story.append(PageBreak())

    # Deep dive tests from mission §2
    story.append(H1("۳. آزمون‌های عمق — پاسخ به سؤال‌های فرایند"))
    story.append(H2("آزمون ۱ — خرید عمده و تأیید پخش"))
    story.append(
        P(
            "پین‌دیس: هدف کاربر تأمین قفسه است. فیلدهای ورود از مایکت و ماده۴ قوانین احراز می‌شود. ماده۷ صریح می‌گوید افزودن به سبد رزرو نیست و تا ثبت نهایی قیمت/موجودی می‌تواند عوض شود — یعنی «ثبت سبد» به سفارش قطعی نمی‌رسد مگر نهایی‌سازی+پرداخت. مقایسه دو پیشنهاد زنده پشت لاگین انجام نشد (مرز دسترسی). بنک مارکت همان هدف را با مالک موجودی واحد و قیمت کارتنی عمومی حل می‌کند (تفاوت مدل: تک‌عرضه‌کننده). بنک‌بازار به استعلام/توافق می‌رسد نه چک‌اوت. مسئول تأیید در پین‌دیس پس از نهایی: پردازش پلتفرم در روز کاری؛ کیفیت با تأمین‌کننده."
        )
    )
    story.append(H2("آزمون ۲ — تخفیف و تغییر سفارش"))
    story.append(
        P(
            "پخش‌یار ادعای تخفیف پلکانی دارد ولی متن پاسخ FAQ در خزش خالی بود → قاعده «کل تعداد یا مازاد» احراز نشد (Gap G2). آیمیس به‌عنوان نرم‌افزار پخش، تخفیف پلکانی خودکار و حداقل/حداکثر کارتن را به‌عنوان قابلیت محصول اعلام می‌کند — بهترین شاهد قاعده‌مند عمومی، هرچند دموی زنده دیده نشد. جشنواره پین‌دیس در لندینگ هست؛ فرمول پله در قوانین عمومی نیست. برای نمونه «یک بسته کمتر روی آستانه» فقط در شبیه‌سازی REF قابل اجرا اعلام می‌شود."
        )
    )
    story.append(H2("آزمون ۳ — احراز و ادامه پس از تأیید (اسنپ‌پی/دیجی‌پی + موجودی)"))
    story.append(
        P(
            "دیجی‌پی پذیرندگی: کدملی+موبایل هم‌نام → تعریف کسب‌وکار/مالک/شبا/شعبه → ابزار QR/بارکد/لینک/درگاه → انتخاب روش پرداخت و شرایط تسویه صنفی → داشبورد تراکنش و تقویم تسویه. خرید BNPL: مصرف سقف مشتری؛ نکول روی تسویه پذیرنده اثر ندارد. برای خرید اعتباری موجودی فروشگاه این مسیر کافی نیست؛ مسیر مستقل بازار = باروک/کسب‌نو (و نوی‌پو برای کشاورزی)."
        )
    )
    story.append(H2("آزمون ۴ — عملیات و ادمین"))
    story.append(
        P(
            "رخداد نقص مدرک: در دیجی‌پی/اسنپ‌پی مسیر بیرونی «رد → ویرایش → بازگشت صف» مستند است؛ نام منوی داخلی صف مشاهده نشد. رخداد عدم حضور تحویل پین‌دیس: مرجوع + خسارت کاربر (قوانین). مغایرت مالی داخلی پلتفرم‌ها: خود پنل مشاهده نشد؛ فقط سرنخ مرچنت/پرتال."
        )
    )
    story.append(H2("آزمون ۵ — اعتبار بانکی و زنجیره‌ای"))
    story.append(
        P(
            "نوی‌پو: ثبت پالیز → شعبه بانک کشاورزی → ضمانت‌نامه بدون آورده نقدی (خبر رسمی) → شارژ کیف گرین‌پی → خرید بازارگاه/پالیزمارکت یا انتقال حلقه → تسویه سررسید. تفاوت با نئوبانک فردی (ویپاد/بانکینو): مصرف مجاز کالایی/زنجیره‌ای است نه تسهیلات نقدی آزاد مصرف‌کننده. کسب‌نو تجارت: همین منطق برای FMCG خرده‌فروش با سقف تا ۵ میلیارد تومان ادعا."
        )
    )
    story.append(PageBreak())

    # Role money table
    story.append(H2("۳.۶ مقایسه نقش‌های مالی (متقاضی تا نکول)"))
    story.append(
        tbl(
            ["محصول", "متقاضی", "مصرف مجاز", "تأمین‌کننده وجه", "بدهکار", "ضامن", "گیرنده تسویه کالا", "مسئول نکول"],
            [
                ["دیجی‌پی BNPL", "فرد", "پذیرندگان شبکه", "دیجی‌پی/شریک", "فرد", "اغلب بدون", "پذیرنده از دیجی‌پی", "دیجی‌پی (FAQ)"],
                ["دیجی‌پی سازمانی", "سازمان برای کارکنان/تدارکات", "پذیرندگان", "دیجی‌پی", "کارمند/سازمان", "سازمان", "پذیرنده", "حسب قرارداد سازمانی"],
                ["تارا/کیپا رفاهی", "سازمان/فرد", "شبکه رفاهی", "پلتفرم/شریک", "کارمند/فرد", "سازمان در مدل رفاهی", "پذیرنده", "سازمان/پلتفرم حسب طرح"],
                ["باروک کسب‌نو", "فروشگاه FMCG", "تأمین‌کنندگان عضو", "بانک تجارت", "فروشگاه", "وثیقه/چک/ضامن", "تأمین‌کننده از بانک/طرح", "بانک می‌پذیرد (ادعا)"],
                ["نوی‌پو", "ذی‌نفع زنجیره کشاورزی", "نهاده/کالای زنجیره", "بانک کشاورزی", "ذی‌نفع", "ضمانت‌نامه بانکی", "تأمین‌کننده زنجیره", "بانک/وثایق طرح"],
                ["اعتبار سنتی پخش (کارا)", "فروشگاه مشتری پخش", "کالای همان پخش", "خود پخش", "فروشگاه", "حساب‌اعتبار پخش", "پخش", "پخش"],
            ],
            [2.3 * cm, 2.2 * cm, 2.4 * cm, 2.3 * cm, 2.0 * cm, 2.2 * cm, 2.6 * cm, 2.2 * cm],
        )
    )
    story.append(Paragraph(fa("جدول پ — مرز مدل روی بدهکار و نکول"), styles["FACaption"]))
    story.append(PageBreak())

    # Scenario bank
    story.append(H1("۴. بانک سناریوها (شناسنامهٔ کامل)"))
    story.append(
        P(
            f"تعداد سناریوهای متمایز در این نسخه: {len(DATA['scenarios'])}. شاخه‌های جزئی داخل همان شناسه مانده‌اند. موارد کم‌شاهد در برگه Gaps فایل اکسل آمده‌اند. فایل قابل‌ویرایش: Iran-B2B-Scenario-Bank-Phase2.xlsx"
        )
    )

    for s in DATA["scenarios"]:
        block = [
            H3(f"{s['id']} — {s['company']} / {s['product']}"),
            P(f"نقش: {s['role']} | هدف: {s['goal']}", "FASmall"),
            P(f"کانال: {s['channel']} | قلمرو: {s['territory']} | اطمینان: {s['confidence']}", "FASmall"),
            P(f"مرز دسترسی: {s['access_boundary']}", "FATag"),
        ]
        step_rows = []
        for st in s["steps"]:
            step_rows.append(
                [
                    str(st["n"]),
                    st["pre"][:80],
                    st["action"][:70],
                    (st["fields"] or "—")[:55],
                    st["system"][:70],
                    st["decision"][:40],
                    f"{st.get('evidence','')[:28]}|{st.get('label','')}",
                ]
            )
        block.append(
            tbl(
                ["گام", "پیش‌وضعیت", "اقدام", "فیلدها", "پاسخ سامانه", "تصمیم", "شاهد"],
                step_rows,
                [0.8 * cm, 2.8 * cm, 2.6 * cm, 2.4 * cm, 2.8 * cm, 2.0 * cm, 3.2 * cm],
            )
        )
        block.append(P(f"نتیجه: {s['outcome']}", "FASmall"))
        block.append(P(f"وجه/ریسک: {s['money']}", "FASmall"))
        block.append(P(f"استثنا: {' | '.join(s.get('exceptions') or ['—'])}", "FASmall"))
        block.append(hr())
        story.append(KeepTogether(block))

    story.append(PageBreak())

    # Financial timeline
    story.append(H1("۵. مدل مالی روی خط زمانی سفارش"))
    story.append(H2("۵.۱ باروک/کسب‌نو (الگوی موجودی سوپر)"))
    story.append(
        tbl(
            ["رویداد", "محرک", "مانده قابل مصرف", "بدهی فروشگاه", "طلب تأمین‌کننده", "وجه واقعی", "ریسک"],
            [
                ["تصویب سقف", "قرارداد شعبه", "مثلاً ۲۰۰م", "۰", "۰", "—", "رد اعتبارسنجی"],
                ["مصرف روی فاکتور", "پرداخت باروکی", "↓ مبلغ فاکتور", "↑ همان", "طلب از بانک/طرح", "کالا می‌رود", "کسری سقف"],
                ["تسویه تأمین‌کننده", "قاعده طرح", "بدون تغییر لزوماً", "پایدار", "↓ با واریز بانک", "بانک→تأمین", "تأخیر عملیاتی"],
                ["بازپرداخت ۱–۶م", "سررسید", "↑ به اندازه اصل (گردان)", "↓", "۰", "حساب فروشگاه→بانک", "نکول→بانک/وثیقه"],
            ],
            [2.6 * cm, 2.4 * cm, 2.4 * cm, 2.4 * cm, 2.6 * cm, 2.4 * cm, 2.0 * cm],
        )
    )
    story.append(Paragraph(fa("جدول ج — خط زمانی کسب‌نو/باروک"), styles["FACaption"]))

    story.append(H2("۵.۲ دیجی‌پی پذیرنده (الگوی مصرف‌کننده — برای ترجمه)"))
    story.append(
        tbl(
            ["رویداد", "بدهکار", "طلب‌دار", "گیرنده تسویه فروش", "اثر نکول مشتری"],
            [
                ["مصرف BNPL", "مشتری", "دیجی‌پی", "—", "—"],
                ["تسویه پذیرنده", "دیجی‌پی", "—", "پذیرنده", "بی‌اثر طبق FAQ"],
                ["قسط مشتری", "مشتری", "دیجی‌پی", "—", "ریسک نزد دیجی‌پی"],
            ],
            [3.5 * cm, 3.0 * cm, 3.0 * cm, 3.5 * cm, 3.5 * cm],
        )
    )

    story.append(H2("۵.۳ آزمون دوپخشی (فرض سناریو / اصلاح عدد فاز۱)"))
    story.append(
        P(
            "جمع پس از تخفیف A: ۱۶۵٫۶م + B: ۷۰م = ۲۳۵٫۶م. سقف ۲۰۰م → کسری اولیه ۳۵٫۶م (نه ۲۵٫۶م). عدد ۲۵٫۶م فقط پس از مرجوعی ۱۰م از A معنا دارد (مانده نسبت به سقف پس از اصلاح: ۲۲۵٫۶−۲۰۰=۲۵٫۶)."
        )
    )
    story.append(
        tbl(
            ["حالت", "اقدام عبور از مانع", "مانده اعتبار", "بدهی", "طلب A", "طلب B", "یادداشت"],
            [
                ["t0 مانع", "نیاز تصمیم", "۲۰۰", "۰", "۱۶۵٫۶", "۷۰", "اگر رزرو کامل لازم باشد گیر می‌کند"],
                ["عبور نقدی", "نقد ۳۵٫۶ + اعتبار ۲۰۰", "۰", "۲۰۰", "۱۶۵٫۶", "۷۰", "فرض"],
                ["عبور کاهش سبد", "حذف ۳۵٫۶ از کالا", "۰ پس از مصرف ۲۰۰", "۲۰۰", "متغیر", "متغیر", "فرض"],
                ["مرجوعی ۱۰م قبل تسویه", "کاهش فاکتور A", "آزادسازی ۱۰ در طراحی ایده‌آل", "۱۹۰ اگر آزاد شود", "۱۵۵٫۶", "۷۰", "شبیه‌سازی"],
                ["مرجوعی ۱۰م بعد تسویه", "بستانکار/سند اصلاحی", "پیچیده", "ممکن است ۲۰۰ بماند تا اصلاح بانک", "طلب خالص↓", "۷۰", "شاهد بازار کامل نیست"],
            ],
            [2.8 * cm, 3.2 * cm, 2.2 * cm, 2.0 * cm, 2.0 * cm, 1.8 * cm, 2.6 * cm],
        )
    )
    story.append(Paragraph(fa("جدول د — شبیه‌سازی دوپخشی (SIM-01)"), styles["FACaption"]))
    story.append(PageBreak())

    # Matrix
    story.append(H1("۶. ماتریس مقایسهٔ بازسازی‌شده (محصول×گام)"))
    story.append(
        tbl(
            ["گام", "پین‌دیس", "بنک‌مارکت", "بنک‌بازار", "پخش‌یار", "باروک", "نوی‌پو", "دیجی‌پی مرچنت"],
            [
                ["احراز کسب‌وکار", "بله", "؟", "جزئی", "ادعا", "بله", "بله", "پذیرنده"],
                ["قیمت کارتنی", "ادعا", "مشهود", "توافقی", "ادعا", "نزد تأمین", "بله", "n/a"],
                ["چند عرضه", "ادعا", "خیر", "بله", "ادعا", "شبکه", "شبکه", "n/a"],
                ["پرداخت اعتبار موجودی", "اشاره+کاربر", "خیرمشهود", "؟", "ادعا", "هسته", "هسته", "خیر"],
                ["تسویه تضمینی فروشنده", "؟", "n/a", "n/a", "؟", "بله ادعا", "بله", "بله FAQ"],
                ["شاهد توقف", "لاگین سبد", "لاگین", "قرارداد حجره", "متن FAQ", "سایت۵۰۰", "لاگین پالیز", "صف ادمین"],
            ],
            [2.4 * cm, 2.0 * cm, 2.0 * cm, 2.0 * cm, 1.8 * cm, 1.8 * cm, 1.8 * cm, 2.4 * cm],
        )
    )
    story.append(Paragraph(fa("جدول ه — وضعیت شاهد؛ جزئیات در XLSX برگه ComparisonMatrix"), styles["FACaption"]))

    # Reference scenario
    story.append(H1("۷. سناریوی مرجع پیشنهادی (طراحی)"))
    story.append(
        P(
            "REF-B2B-CORE-01 و REF-B2B-RETURN-01 مسیر پایه طراحی‌اند (برچسب simulation_design). پایه از پین‌دیس/آیمیس/کارا؛ شاخه اعتبار از باروک/نوی‌پو؛ جداسازی ریسک از دیجی‌پی ترجمه شده به رابطه بانک–پخش."
        )
    )
    story.append(H2("۷.۱ مسیر موفق پایه"))
    for line in [
        "ورود و احراز صنف+لوکیشن+مجوز (الگو: MP-PINDIS-REG-01).",
        "کاتالوگ با واحد کارتن و قیمت مؤثر؛ مقایسه چند پیشنهاد در صورت چندعرضه.",
        "سبد چندپخشی با نمایش سقف اعتبار و کسری؛ رزرو موجودی فقط پس از تأیید پرداخت/اعتبار (الگو: ماده۷ پین‌دیس).",
        "تأیید پخش بر اساس موجودی و سقف اعتبار مشتری نزد پخش (الگو: کارا).",
        "تحویل با رسید دیجیتال؛ ثبت کسری در همان لحظه.",
        "تسویه پخش از محل بانک‌اعتبار؛ بازپرداخت فروشگاه به بانک؛ مانده‌ها جدا در دفتر معین.",
    ]:
        story.append(P("• " + line))
    story.append(H2("۷.۲ شاخه اعتبار"))
    story.append(
        P(
            "محل تصمیم تصویب سقف: بانک/تجمیع‌کننده پیش از یا موازی با اولین سفارش بزرگ. مرز مسئولیت: پخش کالا و کیفیت؛ بانک وجه و نکول؛ پلتفرم تطبیق سفارش و اعلان. حالت بازگشت: لغو قبل مصرف → بدون بدهی؛ مرجوعی قبل تسویه → آزادسازی؛ مرجوعی بعد تسویه → سند اصلاحی سه‌طرفه."
        )
    )
    story.append(H2("۷.۳ معیار سنجش"))
    story.append(
        P(
            "زمان تا اولین سفارش؛ نرخ برگشت پرونده احراز؛ اختلاف تحویل؛ زمان رفع مغایرت تسویه؛ نرخ رها کردن سبد به‌خاطر سقف اعتبار."
        )
    )
    story.append(PageBreak())

    story.append(H1("۸. بهترین رویه‌ها"))
    story.append(
        tbl(
            ["رویه", "مبدأ", "شاهد", "شرط انتقال به سوپر", "برچسب"],
            [
                ["اعتبار موجودی بانک‌محور + تضمین تأمین‌کننده", "CR-BAROOK-BUY-01", "E-ECOTAM-01", "شبکه تأمین FMCG عضو طرح", "الگوی مشاهده‌شده"],
                ["سبد غیررزروی تا نهایی", "MP-PINDIS-ORDER-01", "E-PD-TERMS-01", "همه بازارگاه‌ها", "مستند رسمی"],
                ["تسویه فروشنده ≠ وصول اقساط مشتری", "CR-DIGIPAY-C-BNPL-01", "E-DP-02", "ترجمه به بانک–پخش", "مستند رسمی"],
                ["شفافیت مالک کیفیت/فاکتور", "MP-PINDIS-DELIVERY-01", "E-PD-TERMS-01", "چندعرضه‌کننده", "مستند رسمی"],
                ["نمایش کسری سقف قبل تأیید", "FIN-TWOSUPPLIER-SIM-01", "SIM-01", "هر B2B اعتباری", "طراحی"],
                ["مرزبندی رفاه از موجودی", "CR-DIGIPAY-ORG-01", "E-DP-B2O-01", "حکمرانی محصول", "تحلیلی"],
                ["مرجوعی قبل/بعد تسویه جدا", "REF-B2B-RETURN-01", "SIM+Gap", "مقیاس‌پذیری", "طراحی"],
                ["ویزیتوری آفلاین مکمل", "SW-NANO-01", "E-NANO-01", "نفوذ پخش سنتی", "ادعای محصول"],
            ],
            [4.2 * cm, 3.0 * cm, 2.6 * cm, 3.4 * cm, 2.4 * cm],
        )
    )
    story.append(Paragraph(fa("جدول و — رویه‌ها؛ جزئیات در XLSX/BestPractices"), styles["FACaption"]))

    story.append(H1("۹. خلأها و اقدام پژوهشی بعدی"))
    story.append(
        tbl(
            ["ID", "خلأ", "اقدام بعدی"],
            [
                ["G1", "ادغام رسمی پین‌دیس×باروک", "حساب آزمایشی یا مصاحبه تأمین‌کننده عضو هر دو"],
                ["G2", "قاعده پله تخفیف پخش‌یار", "دمو یا اسکرین قوانین قیمت"],
                ["G3", "کسری تحویل پین‌دیس", "راهنمای مرجوعی داخل اپ"],
                ["G4", "کارمزد دقیق کسب‌نو", "استعلام شعبه/قرارداد"],
                ["G5", "پنل ادمین داخلی", "دمو رسمی نقش عملیات"],
                ["G6", "جیحون و آرین تدبیر", "اسناد کارفرما"],
                ["G7", "اسنپ‌پی×ساپلای همان سفارش", "تست حساب رستوران"],
                ["G8", "مرجع کیفیت بنک‌بازار", "قوانین حجره"],
            ],
            [1.5 * cm, 5.5 * cm, 9.5 * cm],
        )
    )
    story.append(PageBreak())

    story.append(H1("۱۰. پیوست شواهد و فایل‌های همراه"))
    story.append(P("شناسه‌ها میان PDF و Iran-B2B-Scenario-Bank-Phase2.xlsx یکسان‌اند. JSON خام: scenarios_phase2.json"))
    for label, url in [
        ("پین‌دیس قوانین", "https://pindis.com/terms-and-conditions/"),
        ("مایکت پین‌دیس", "https://myket.ir/app/com.pindis"),
        ("اکوتام کسب‌نو", "https://ecotam.ir/"),
        ("گرین‌پی", "https://greenpay360.ir/"),
        ("دیجی‌پی FAQ", "https://www.mydigipay.com/academy/faq/"),
        ("دیجی‌پی B2O", "https://www.mydigipay.com/landing/B2O/"),
        ("دیجی‌پی آکادمی ثبت‌نام", "https://www.mydigipay.com/academy/"),
        ("اسنپ‌پی آکادمی پذیرندگی", "https://academy.snapppay.ir/merchant-acquisition/"),
        ("آیمیس عمده", "https://imis.ir/solutions/wholesale-online/"),
        ("بنک مارکت", "https://bonak.market/"),
        ("بنک‌بازار", "https://www.bonakbazar.com/"),
        ("اسنپ ساپلای", "https://snapp.supply/"),
        ("پخش‌یار", "https://pakhshiar.ir/"),
        ("تارا سازمانی", "https://taracredit.com/org-req/"),
    ]:
        story.append(Paragraph(f"{fa(label)}: {link(url)}", styles["FASmall"]))

    story.append(Spacer(1, 0.5 * cm))
    story.append(hr())
    story.append(
        P(
            "معیار پایان مأموریت: خواننده از هر تصمیم پیشنهادی به شناسهٔ سناریو و شاهد می‌رسد؛ در هر گام می‌داند چه داده وارد می‌شود، چه نقشی تصمیم می‌گیرد، و مسیر کجا به مرز دسترسی یا خلأ می‌رسد. بسته تحویل = PDF + XLSX + JSON."
        )
    )

    out = ROOT / "Iran-B2B-Credit-Competitor-Benchmark-FA.pdf"
    doc = SimpleDocTemplate(
        str(out),
        pagesize=A4,
        rightMargin=1.4 * cm,
        leftMargin=1.4 * cm,
        topMargin=1.5 * cm,
        bottomMargin=1.7 * cm,
        title="فاز۲ بررسی رقبا خرید عمده و اعتبار ایران",
        author="Cursor Research Agent",
    )
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    # also copy phase2-named artifact
    phase2_name = ART / "Iran-B2B-Journey-Scenarios-Phase2-FA.pdf"
    import shutil

    shutil.copy2(out, ART / "Iran-B2B-Credit-Competitor-Benchmark-FA.pdf")
    shutil.copy2(out, phase2_name)
    shutil.copy2(out, ROOT / "Iran-B2B-Journey-Scenarios-Phase2-FA.pdf")
    print("PDF", out, "pages built; size", out.stat().st_size)
    return out


if __name__ == "__main__":
    build_xlsx()
    build_pdf()
