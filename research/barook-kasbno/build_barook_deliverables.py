# -*- coding: utf-8 -*-
"""Barook / Kasbno deep research deliverables — auditable Excel + RTL HTML/PDF + Scenario 0.3."""
from __future__ import annotations

import json
import subprocess
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
ART = Path(r"D:\HomeDesign\artifacts")
ART.mkdir(exist_ok=True)
OBS = "1405/09/03"

HF = PatternFill("solid", fgColor="0B3D5C")
HFont = Font(color="FFFFFF", bold=True, name="Tahoma", size=10)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="right")
THIN = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)


def sheet(ws, headers, rows, widths=None):
    ws.append(headers)
    for i in range(1, len(headers) + 1):
        c = ws.cell(1, i)
        c.fill = HF
        c.font = HFont
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    for r in rows:
        ws.append(list(r))
        for i in range(1, len(headers) + 1):
            cell = ws.cell(ws.max_row, i)
            cell.alignment = WRAP
            cell.border = THIN
            cell.font = Font(name="Tahoma", size=9)
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ── Evidence ───────────────────────────────────────────────────────────────
EV = [
    ["EV-BK-001", "https://barook.ir/fmcg", OBS, "صفحه رسمی اعتبار کسب‌وکاری سوپرمارکت", "سقف تا ۵میلیارد تومان؛ کارمزد بانک ۲٫۲۵٪ سالانه از روز خرید؛ چک ۱۲۵٪ برای A/B؛ گام‌های ثبت تا بازپرداخت؛ لیست ۱۰ پخش؛ کارمزد باروک در ماشین‌حساب نمونه", "official_product", "A"],
    ["EV-BK-002", "https://barook.ir/", OBS, "خانه باروک", "اعتبار کسب‌وکاری/مصرف‌کننده/شرکتی؛ پشتوانه تجارت و ملت؛ تماس ۰۲۱۹۱۰۰۳۹۹۵", "official_landing", "A"],
    ["EV-BK-003", "https://cafebazaar.ir/app/ir.barook.barookcustomer", OBS, "اپ بازار — باروک", "بسته ir.barook.barookcustomer؛ ناشر بامداد رویان کیان؛ ۵۲هزار نصب؛ به‌روز ۱۴۰۵/۰۶/۱۴؛ بازارگاه اعتباری اصناف و حقیقی", "app_store", "A"],
    ["EV-BK-004", "https://ecotam.ir/", OBS, "اکوتام — محصول کسب‌نو", "کسب‌نو FMCG با همکاری باروک؛ سقف ۵۰ میلیارد ریال با کمترین کارمزد؛ جدا از تأمین‌نو/صنف‌نو/کشت‌نو", "official_scf", "A"],
    ["EV-BK-005", "https://tejaratbank.ir/", OBS, "فهرست طرح‌های بانک تجارت", "کارت کسب‌نو در لیست طرح‌ها؛ صفحه جزئیات مستقیم در خزش SSL مشکل داشت", "bank_portal_index", "B"],
    ["EV-BK-006", "https://purson.ir/fa/content/931240/...", OBS, "بازنشر پرسون با انتساب بانک", "مراحل ثبت؛ سقف ۵۰ میلیارد ریال؛ بازپرداخت ۱–۶ ماه یکجا؛ شارژ به میزان اصل قسط؛ نقش باروک تجمیع‌کننده", "secondary_bank_echo", "B"],
    ["EV-BK-007", "https://www.sharghdaily.com/.../1013488-...", OBS, "شرق — معرفی وام کسب‌نو", "۵ میلیارد تومان؛ بدون نیاز به چک یا ضمانت‌نامه در متن خبر (تعارض با صفحه رسمی چک ۱۲۵٪)", "secondary_news", "C"],
    ["EV-BK-008", "https://way2pay.ir/طرح-کسب‌نو؛-پیشنهاد-بانک-تجارت-برای-خ/", OBS, "راه‌پرداخت ۲۳ فروردین ۱۴۰۴", "همان ارقام پرسون؛ همکاری باروک", "secondary_fintech_media", "B"],
    ["EV-BK-009", "https://hugco.ir/services/credit/", OBS, "هاگ — راهنمای خرید اعتباری باروک‌کارت", "۳۰/۶۰/۹۰؛ ۲٫۲۵٪؛ کشیدن کارت روی پوز موزع؛ لیست پذیرندگان شامل سایه‌سمن/گلرنگ/به پخش؛ مرجوعی؛ سفارش آنلاین+ویزیت", "partner_official_guide", "A"],
    ["EV-BK-010", "https://sapegah.com/barook/", OBS, "ساپ — یکپارچگی ERP با باروک", "چک ۱۲۵٪؛ کارمزد ۲٪ (متعارض با ۲٫۲۵٪ رسمی)؛ جریان فاکتور→باروک→بانک", "integrator_secondary", "C"],
    ["EV-BK-011", "https://telegram.me/s/barook_ir?q=%23آموزش", OBS, "کانال رسمی آموزش‌ها", "فعالسازی کارت؛ ثبت اعتبار؛ POS P10 توسن‌تکنو؛ پرداخت نشده‌ها روی پوز؛ آپلود فاکتور تأمین‌کننده؛ شارژ کیف؛ حقوقی با/بدون توکن", "official_telegram_trainings", "A"],
    ["EV-BK-012", "https://smtnews.ir/.../84558-...", OBS, "صمت — کیف اعتباری سایه‌سمن", "۵۰ میلیون بدون اعتبارسنجی برای ۴۰هزار مشتری؛ ریسک نزد سایه‌سمن؛ افزایش سقف→بانک", "partner_news", "B"],
    ["EV-BK-013", "https://jahaneghtesad.com/.../189977-...", OBS, "جهان اقتصاد — سفره اعتباری سن‌ایچ/سایه‌سمن", "تأیید مدل ریسک‌پذیر پخش؛ شهریور ۱۴۰۵", "partner_news", "B"],
    ["EV-BK-014", "https://khabarfarsi.com/u/261722072", OBS, "دیجیاتو via خبر فارسی — اسنپ‌مارکت×باروک", "سوپرهای اسنپ‌مارکت به شبکه اعتباری می‌پیوندند؛ خرج در مراکز از جمله اسنپ‌مارکت پرو؛ ۱۴۰۵/۰۷/۲۹", "partner_announce", "B"],
    ["EV-BK-015", "https://pro.snapp.market/", OBS, "ورودی اسنپ‌مارکت پرو", "صفحه عمومی کم‌محتوا در خزش؛ راهنمای سفارش اعتباری باروک روی این دامنه دیده نشد", "partner_portal_thin", "D"],
    ["EV-BK-016", "https://app.barook.ir/", OBS, "وب‌اپ باروک", "HTTP ۲۰۰؛ پوسته ورود؛ جزئیات پس از لاگین مشاهده نشد (بدون حساب پژوهشی)", "app_entry_behind_auth", "B"],
    ["EV-BK-017", "https://way2pay.ir/credit-supply-chain-growth-engine/", OBS, "نشست اعتبار موتور رشد — باروک", "خدمات مالی تعبیه‌شده در نقطه سفارش پخش", "event_coverage", "C"],
    ["EV-BK-018", "phase1/2 repo evidence E-ECOTAM/E-BAROOK", OBS, "خط مبنای پژوهش قبلی HomeDesign", "الگوی موجودی+اعتبار؛ gap کارمزد ۲٪", "prior_research", "B"],
]

ENTITIES = [
    ["ENT-BAROOK", "platform", "باروک", "بامداد رویان کیان", "بازارگاه/تجمیع‌کننده اعتباری؛ صدور کارت و گزارش", "barook.ir"],
    ["ENT-TEJARAT", "bank", "بانک تجارت", "بانک تجارت", "مصوب‌کننده حد؛ کارمزد؛ وصول خودکار از حساب جاری", "tejaratbank.ir"],
    ["ENT-ECOTAM", "scf_portal", "اکوتام", "SCF بانک تجارت", "چتر محصولات؛ کسب‌نو یکی از شش محصول", "ecotam.ir"],
    ["ENT-MELLAT", "bank", "بانک ملت", "بانک ملت", "در فوتر باروک به‌عنوان پشتوانه برخی طرح‌ها نام برده شده — نقش در کسب‌نو FMCG روشن نیست", "barook.ir footer"],
    ["ENT-HUG", "distributor_partner", "هاگ", "هاگ", "پذیرنده + راهنمای عملیاتی کشیدن کارت", "hugco.ir"],
    ["ENT-SAYESAMAN", "distributor", "سایه سمن / سن‌ایچ", "پخش سایه سمن", "پذیرنده + کیف اختصاصی ریسک‌پذیر پخش", "news 1405"],
    ["ENT-SNAPPMKT", "marketplace", "اسنپ‌مارکت / پرو", "اسنپ", "اعلام همکاری اعتباری؛ پرو به‌عنوان مرکز خرید", "digiato 1405/07/29"],
    ["ENT-GOLRANG", "distributor", "گلرنگ پخش", "", "نام در لیست پذیرندگان باروک FMCG", "barook.ir/fmcg"],
    ["ENT-BEHPAKHSH", "distributor", "به پخش", "", "نام در لیست پذیرندگان", "barook.ir/fmcg"],
    ["ENT-GOLSTAN", "distributor", "گلستان", "", "نام در لیست پذیرندگان", "barook.ir/fmcg"],
    ["ENT-TABIAT", "distributor", "بخش سراسری طبیعت", "", "نام در لیست پذیرندگان", "barook.ir/fmcg"],
    ["ENT-HIMART", "distributor", "هایمارت", "", "نام در لیست پذیرندگان", "barook.ir/fmcg"],
    ["ENT-ALSA", "distributor", "السا پخش", "", "نام در لیست پذیرندگان", "barook.ir/fmcg"],
    ["ENT-KHANSAR", "distributor", "پخش خوانسار", "", "نام در لیست پذیرندگان", "barook.ir/fmcg"],
    ["ENT-GOLPAKHSH1", "distributor", "گل پخش اول", "", "نام در لیست پذیرندگان", "barook.ir/fmcg"],
    ["ENT-AZIMI", "distributor", "پخش عظیمی", "", "نام در لیست پذیرندگان", "barook.ir/fmcg"],
    ["ENT-TOSAN", "pos_vendor", "توسن تکنو P10", "", "پوز اندرویدی شرکای باروک طبق آموزش تلگرام", "telegram training"],
]

PRODUCTS = [
    ["PRD-KASBNO", "کسب‌نو", "Tejarat+Barook", "FMCG supermarket credit", "active_public", "EV-BK-001;EV-BK-004"],
    ["PRD-BAROOKCARD", "باروک‌کارت", "Barook", "physical/digital card for spend at contracted acceptors", "active_public", "EV-BK-001;EV-BK-003"],
    ["PRD-SAYE-WALLET", "کیف اعتباری اختصاصی سایه‌سمن", "Barook+SayeSaman", "50m without bank scoring; risk on distributor", "announced_ops", "EV-BK-012"],
    ["PRD-ROOZNO", "روز نو", "Tejarat Ecotam", "consumer SCF adjacent", "adjacent", "EV-BK-004"],
    ["PRD-TAMINO", "تأمین نو", "Tejarat Ecotam", "production SCF / electronic draft — NOT Kasbno", "adjacent_do_not_conflate", "EV-BK-004"],
    ["PRD-SENFNO", "صنف نو / پلاس", "Tejarat Ecotam", "distribution SCF legal entities", "adjacent", "EV-BK-004"],
    ["PRD-KESHTNO", "کشت نو", "Tejarat Ecotam", "agri SCF", "adjacent", "EV-BK-004"],
    ["PRD-DIGITAL", "اعتبار کسب‌وکاری کالای دیجیتال", "Barook", "non-FMCG vertical", "active_public", "EV-BK-002"],
    ["PRD-CORP", "اعتبار شرکتی / سازمانی / درمان", "Barook", "other personas", "active_public", "EV-BK-002"],
]

RELATIONS = [
    ["REL-BK-01", "ENT-TEJARAT", "ENT-BAROOK", "product_partner_kasbno", "تجمیع‌کننده ثبت‌نام/کارت/گزارش", 2, "EV-BK-001;EV-BK-004", "NEW/CONFIRM official"],
    ["REL-BK-02", "ENT-ECOTAM", "PRD-KASBNO", "hosts_product_page", "کسب‌نو در اکوتام", 2, "EV-BK-004", "CONFIRM"],
    ["REL-BK-03", "ENT-BAROOK", "ENT-HUG", "acceptor_with_ops_guide", "خرید+پوز+سررسید", 3, "EV-BK-009", "NEW depth vs prior"],
    ["REL-BK-04", "ENT-BAROOK", "ENT-SAYESAMAN", "acceptor_plus_private_wallet", "لیست FMCG + کیف اختصاصی", 2, "EV-BK-001;EV-BK-012", "NEW wallet model"],
    ["REL-BK-05", "ENT-BAROOK", "ENT-SNAPPMKT", "announced_credit_network", "اعلام همکاری؛ خرج در پرو ادعا", 1, "EV-BK-014", "NEW announce L1"],
    ["REL-BK-06", "ENT-BAROOK", "ENT-GOLRANG", "listed_acceptor", "نام در FMCG", 1, "EV-BK-001", "CONFIRM name"],
    ["REL-BK-07", "ENT-BAROOK", "ENT-BEHPAKHSH", "listed_acceptor", "نام در FMCG", 1, "EV-BK-001", "CONFIRM"],
    ["REL-BK-08", "ENT-BAROOK", "ENT-TOSAN", "pos_hardware", "P10 آموزش", 2, "EV-BK-011", "NEW"],
    ["REL-BK-09", "ENT-BAROOK", "ENT-MELLAT", "mentioned_backing", "فوتر؛ نقش کسب‌نو نامشخص", 0, "EV-BK-002", "lead"],
    ["REL-BK-10", "PRD-SAYE-WALLET", "ENT-TEJARAT", "risk_transfer_on_uplift", "افزایش سقف→اعتبارسنجی بانک", 1, "EV-BK-012", "NEW"],
]

CLAIMS = [
    ["CL-01", "PRD-KASBNO", "supermarket", "ceiling", "تا ۵۰ میلیارد ریال / ۵ میلیارد تومان", "1405", "EV-BK-001;EV-BK-004", "A", "سقف تبلیغاتی؛ سقف واقعی پس از بانک"],
    ["CL-02", "PRD-KASBNO", "supermarket", "bank_fee", "۲٫۲۵٪ سالانه روزشمار از روز خرید", "1405", "EV-BK-001", "A", "حل تعارض با ۲٪ منابع قدیمی"],
    ["CL-03", "PRD-KASBNO", "supermarket", "barook_fee_sample", "کارمزد باروک ۵۰۰٬۰۰۰ تومان در نمونه ماشین‌حساب ۱۰۰م", "1405", "EV-BK-001", "B", "نیاز به قاعده عمومی تعرفه"],
    ["CL-04", "PRD-KASBNO", "supermarket", "guarantee_cheque", "یک برگه چک ۱۲۵٪ اعتبار برای رتبه A/B؛ C نیاز ضامن", "1405", "EV-BK-001", "A", "تعارض با خبر «بدون چک»"],
    ["CL-05", "PRD-KASBNO", "supermarket", "repay_window", "بازپرداخت خرید ۱–۶ ماه یکجا (اکوتام/خبر) و ۳۰/۶۰/۹۰ در هاگ", "mixed", "EV-BK-004;EV-BK-009", "B", "بازه رسمی بانک vs انتخاب سررسید پذیرنده"],
    ["CL-06", "PRD-KASBNO", "supermarket", "revolving", "شارژ مجدد پس از بازپرداخت؛ پرسون: به میزان اصل قسط", "echo", "EV-BK-006", "B", "تعارض «معادل قسط» vs «اصل قسط»"],
    ["CL-07", "PRD-KASBNO", "supplier", "default_risk", "بانک نکول را می‌پذیرد / پخش در سررسید از بانک می‌گیرد", "1405", "EV-BK-001;EV-BK-009", "A", "برای فاکتور باروکی بانک‌محور"],
    ["CL-08", "PRD-KASBNO", "supplier", "advance_50pct", "تأمین تا ۵۰٪ منابع تأمین‌کننده / تنزیل نیمی با سود ۲۳٪", "mixed", "EV-BK-006;EV-BK-001", "B", "دو بیان متفاوت"],
    ["CL-09", "PRD-SAYE-WALLET", "saye_customers", "no_bank_score_50m", "۵۰م بدون اعتبارسنجی؛ ریسک پخش", "1405/06", "EV-BK-012", "B", "جدا از مدل بانک‌محور"],
    ["CL-10", "REL-BK-05", "snapp_sellers", "join_network", "پیوستن سوپرهای اسنپ‌مارکت به شبکه؛ خرج در پرو", "1405/07/29", "EV-BK-014", "B", "فرم خرید مشاهده نشد"],
]

ORDER_STEPS = [
    ["OS-BK-01", "B2B-01", "onboarding", "فروشگاه", "barook.ir / app", "ثبت‌نام؛ هویت؛ فروشگاه؛ کارت ملی؛ سند کسب", "کد ملی+موبایل اعتبارسنجی اولیه", "documented", "EV-BK-001", "", ""],
    ["OS-BK-02", "B2B-01", "branch", "فروشگاه+بانک", "شعبه تجارت منتخب", "مراجعه؛ حساب جاری؛ چک ضمانت؛ قرارداد", "تخصیص سقف؛ ارسال کارت", "documented", "EV-BK-001;EV-BK-009", "", ""],
    ["OS-BK-03", "B2B-02/05", "order_channel_visit", "فروشگاه+ویزیتور", "ویزیت پخش طرف قرارداد", "اعلام داشتن باروک‌کارت؛ مذاکره سررسید", "فاکتور با سررسید", "documented_partner", "EV-BK-001;EV-BK-009", "", "کاتالوگ باروک‌ترید چندپخشی دیده نشد"],
    ["OS-BK-04", "B2B-02/05", "order_channel_online_hug", "فروشگاه", "سفارش آنلاین هاگ", "انتخاب روش باروک‌کارت", "پرداخت هنگام تحویل با کارت", "documented_partner", "EV-BK-009", "", ""],
    ["OS-BK-05", "B2B-06/PAY-CREDIT", "spend", "فروشگاه+موزع", "پوز باروک روی تحویل", "کشیدن باروک‌کارت؛ سررسید از قبل روی فاکتور", "مصرف اعتبار؛ کالا تحویل", "documented", "EV-BK-001;EV-BK-009;EV-BK-011", "anon-txn-not-captured", "تراکنش زنده پژوهشی ثبت نشد"],
    ["OS-BK-06", "B2B-06", "pos_unpaid_list", "موزع", "اپ باروک روی POS", "سفارشات پرداخت‌نشده → پرداخت → کشیدن کارت", "تسویه پیش‌فاکتور روی پوز", "training_video", "EV-BK-011", "", "دقیقه دقیق ویدئو در ویجت تلگرام کامل استخراج نشد"],
    ["OS-BK-07", "B2B-07", "supplier_invoice_upload", "تأمین‌کننده", "پنل تأمین‌کننده", "آپلود فاکتورها", "ثبت فاکتور در سامانه", "training_video", "EV-BK-011", "", ""],
    ["OS-BK-08", "B2B-09", "repay", "بانک+فروشگاه", "حساب جاری تجارت", "واریز قبل سررسید؛ کسر خودکار اصل+کارمزد ۲٫۲۵٪", "شارژ مجدد مانده اعتبار", "documented", "EV-BK-001", "", "زمان دقیق برگشت سقف در لاگ سامانه دیده نشد"],
    ["OS-BK-09", "D-09", "early_settle", "فروشگاه", "درخواست تسویه زودهنگام در باروک (هاگ)", "کارمزد روزشمار کمتر؛ برگشت فوری به سقف (ادعای هاگ)", "—", "partner_claim", "EV-BK-009", "", "نیاز تأیید باروک/بانک"],
    ["OS-BK-10", "D-06", "return_before_due", "فروشگاه+پخش", "مرجوعی هاگ", "کاهش مبلغ بازپرداخت سررسید اگر قبل سررسید ثبت شود", "—", "partner_claim", "EV-BK-009", "", "پس از سررسید → اعتبار حساب فروشگاه"],
]

PAY = [
    ["PAY-CREDIT-KASBNO", "credit", "BarookCard spend", "Tejarat current account auto-debit on due", "at card swipe / invoice due set with distributor", "EV-BK-001", "bank fee 2.25% p.a. day-count; Barook fee separate sample"],
    ["PAY-CASH-WALLET", "cash_wallet", "شارژ کیف پول اپ", "درگاه داخل اپ", "for cash purchases in app ecosystem", "EV-BK-011", "not same as Kasbno credit line"],
    ["PAY-GUARANTEE-CHEQUE", "guarantee_not_purchase", "چک ضمانت خط اعتبار", "بانک/طرح", "before card issuance", "EV-BK-001", "NOT PAY-CHEQUE for goods"],
    ["PAY-SAYE-WALLET", "distributor_credit", "کیف سایه‌سمن", "سایه‌سمن ریسک", "spend in Barook network claim", "EV-BK-012", "distinct from bank Kasbno"],
]

CONFLICTS = [
    ["CF-01", "fee_2_vs_225", "sapegah/older: 2% | barook.ir/fmcg+hug: 2.25%", "RESOLVED_for_product_page", "Use 2.25% as current public Barook FMCG; treat 2% as outdated/secondary", "EV-BK-001 vs EV-BK-010"],
    ["CF-02", "no_cheque_marketing_vs_125pct", "sharghdaily: بدون چک/ضمانت‌نامه | official: چک ۱۲۵٪", "OPEN_wording", "Marketing omits guarantee cheque; official FAQ requires it", "EV-BK-007 vs EV-BK-001"],
    ["CF-03", "repay_1to6m_vs_30_60_90", "Ecotam/news 1–6 months lump | Hug 30/60/90 selectable", "PARTIAL", "Bank product allows up to 6m; acceptor UX may expose 30/60/90", "EV-BK-004 vs EV-BK-009"],
    ["CF-04", "principal_vs_installment_reload", "پرسون: شارژ به میزان اصل قسط | هاگ: همان مبلغ پس از بازپرداخت", "OPEN", "Need bank/Barook definition of reload base (principal only vs principal+fee)", "EV-BK-006 vs EV-BK-009"],
    ["CF-05", "barook_fee_opacity", "Calculator shows 500k Barook fee on 100m sample; FAQ silent on formula", "OPEN", "Ask Barook tariff schedule", "EV-BK-001"],
    ["CF-06", "mellat_role", "Footer cites Mellat+Tejarat; Kasbno texts cite Tejarat only", "OPEN", "Separate which products use Mellat", "EV-BK-002"],
    ["CF-07", "snapp_ops_depth", "Announce join network / spend at Pro | Pro page no Barook checkout observed", "OPEN", "Need guide or test account on Pro", "EV-BK-014 vs EV-BK-015"],
    ["CF-08", "taminnu_instruments_vs_kasbno", "Ecotam general flow uses draft/instruction-payment; Kasbno consumer path is card+POS", "RESOLVED_do_not_merge", "Do not attribute draft mechanics to supermarket card spend without Kasbno-specific doc", "EV-BK-004"],
]

TIMELINE = [
    ["1402/11", "رونمایی کسب‌نو+باروک (لینکدین باروک)", "launch", "EV-BK-008 context"],
    ["1404/01/23", "راه پرداخت معرفی کسب‌نو", "media", "EV-BK-008"],
    ["1404/03/05", "شرق سقف ۵ میلیارد تومان", "media", "EV-BK-007"],
    ["1405/06", "کیف اختصاصی سایه‌سمن/سن‌ایچ", "partner_wallet", "EV-BK-012"],
    ["1405/06/14", "به‌روزرسانی اپ بازار", "app", "EV-BK-003"],
    ["1405/07/29", "اعلام همکاری اسنپ‌مارکت×باروک", "partner", "EV-BK-014"],
    ["1405/09/03", "این پژوهش؛ مشاهده barook.ir/fmcg زنده", "research", "EV-BK-001"],
]

GAPS = [
    ["G-BK-01", "P0", "حساب پژوهشی / سند بی‌نام یک سفارش کامل", "هیچ شناسه تراکنش از سقف→مصرف→تسویه پخش→برگشت سقف نداریم", "Barook ops", "masked invoice+settlement sample", "blocks L4"],
    ["G-BK-02", "P0", "تعرفه کارمزد باروک", "فقط نمونه ماشین‌حساب", "Barook finance", "fee schedule PDF", "MVP pricing"],
    ["G-BK-03", "P0", "تعریف شارژ مجدد اصل vs قسط", "تعارض متون", "Tejarat Kasbno product owner", "contract clause", "D-09"],
    ["G-BK-04", "P1", "اسنپ‌مارکت پرو checkout اعتباری", "اعلام بدون UI", "Snapp+Barook", "screen guide", "partner path"],
    ["G-BK-05", "P1", "API ارجاع vs اتصال مستقیم باروک‌ترید", "مستند عمومی ناکافی", "Barook Trade eng", "API/BRD excerpt", "architecture"],
    ["G-BK-06", "P1", "دقایق ویدئو تلگرام ۱۹/۲۱/۲۹/۵۲", "ویجت بدون متن کامل", "Barook content", "rehost or transcript", "training evidence"],
    ["G-BK-07", "P2", "نقش بانک ملت در FMCG", "فقط فوتر", "Barook", "product matrix", "entity map"],
]

CRAWL = [
    ["CL-001", 1, "start FMCG page", "barook.ir/fmcg", "https://barook.ir/fmcg", "ok_200", "full public steps+fee+acceptors", "EV-BK-001", "app+ecotam"],
    ["CL-002", 1, "WebFetch tool", "barook.ir/fmcg", "WebFetch", "500_tool", "tool failed; urllib succeeded", "", "use urllib"],
    ["CL-003", 1, "app store", "ir.barook.barookcustomer", "https://cafebazaar.ir/app/ir.barook.barookcustomer", "ok", "legal name Bamdad Royan Kian", "EV-BK-003", ""],
    ["CL-004", 1, "ecotam kasbno", "کسب‌نو", "https://ecotam.ir/", "ok", "product card with Barook", "EV-BK-004", ""],
    ["CL-005", 1, "telegram trainings", "#آموزش", "https://telegram.me/s/barook_ir?q=%23آموزش", "ok_partial", "POS unpaid, invoice upload, legal token", "EV-BK-011", "fetch individual posts"],
    ["CL-006", 2, "telegram post 52", "settlement training", "https://t.me/barook_ir/52", "widget_thin", "channel shell only", "", "use search index snippets"],
    ["CL-007", 1, "hug credit guide", "باروک‌کارت هاگ", "https://hugco.ir/services/credit/", "ok", "30/60/90 POS delivery flow", "EV-BK-009", ""],
    ["CL-008", 1, "saye saman wallet", "سایه سمن باروک", "smtnews/jahaneghtesad", "ok", "50m no scoring risk on distributor", "EV-BK-012", ""],
    ["CL-009", 1, "snapp announce", "اسنپ مارکت باروک", "https://khabarfarsi.com/u/261722072", "ok", "network join + Pro spend claim", "EV-BK-014", "pro.snapp.market"],
    ["CL-010", 2, "snapp pro portal", "pro checkout barook", "https://pro.snapp.market/", "thin_page", "no barook payment UI public", "EV-BK-015", "Gap G-BK-04"],
    ["CL-011", 2, "fee conflict resolve", "2% vs 2.25%", "barook.ir/fmcg FAQ", "ok", "official 2.25%", "EV-BK-001", "CF-01 resolved"],
    ["CL-012", 2, "cheque conflict", "بدون چک vs 125%", "shargh vs fmcg", "conflict_logged", "official requires cheque", "CF-02", ""],
    ["CL-013", 1, "app.barook.ir", "login shell", "https://app.barook.ir/", "behind_auth", "no catalog without account", "EV-BK-016", "research account request"],
    ["CL-014", 2, "tejarat SSL probe", "official bank page", "tejaratbank.ir", "ssl_error_local", "index via search; detail page not mirrored", "EV-BK-005", "retry env"],
    ["CL-015", 2, "prior phase link", "kasbno baseline", "research/iran-b2b-benchmark", "ok", "reuse labels; deepen fees/partners", "EV-BK-018", ""],
]

SCENARIO_DELTA = [
    ["B2B-01", "confirm_refine", "ثبت‌نام باروک + مدارک + شعبه تجارت + چک ضمانت قبل کارت", "EV-BK-001;EV-BK-009", "تأیید با اصلاح: چک ضمانت الزامی در FAQ رسمی"],
    ["B2B-02", "refine", "بازار عمده باروک‌ترید چندپخشی هنوز شاهد UI ندارد؛ مسیر واقعی: ویزیتور/سفارش پذیرنده طرف قرارداد", "EV-BK-001;EV-BK-009", "اصلاح فرض بازارگاه چندعرضه تا شاهد UI"],
    ["B2B-03", "insufficient", "سبد چندپخشی باروک‌ترید مشاهده نشد", "", "شاهد ناکافی"],
    ["B2B-04", "insufficient", "پنل پاسخ موجودی پخش در باروک‌ترید دیده نشد؛ آپلود فاکتور تأمین‌کننده در آموزش هست", "EV-BK-011", "شاهد ناکافی برای پیشنهاد اصلاحی"],
    ["B2B-05", "refine", "مبلغ/سررسید روی فاکتور پخش قبل کشیدن کارت قطعی می‌شود (هاگ)", "EV-BK-009", "تأیید جزئی مسیر پذیرنده"],
    ["B2B-06", "confirm", "PAY-CREDIT = کشیدن باروک‌کارت روی پوز باروک موزع؛ سررسید از قبل انتخاب", "EV-BK-001;EV-BK-009;EV-BK-011", "تأیید"],
    ["PAY-CASH", "refine", "شارژ کیف پول اپ مسیر جدا؛ نقد کالایی لزوماً همان کسب‌نو نیست", "EV-BK-011", "اصلاح تفکیک"],
    ["PAY-CHEQUE", "refine", "چک در کسب‌نو عمدتاً ضمانت خط است نه ابزار خرید کالا", "EV-BK-001", "اصلاح"],
    ["PAY-CREDIT", "confirm", "مصرف سقف مشترک بین پذیرندگان؛ بانک کسر خودکار", "EV-BK-001;EV-BK-009", "تأیید"],
    ["B2B-07", "partial", "آپلود فاکتور تأمین‌کننده آموزش داده شده؛ شرط ارسال=پس از کشیدن کارت در مسیر کارت‌خوان", "EV-BK-011;EV-BK-001", "تأیید جزئی"],
    ["B2B-08", "partner_only", "مرجوعی در راهنمای هاگ آمده؛ عمومی باروک کامل نیست", "EV-BK-009", "شاهد ناکافی عمومی"],
    ["B2B-09", "confirm_partial", "تسویه بانک از حساب جاری؛ تسویه پخش از بانک در سررسید ادعا؛ سند همان سفارش نداریم", "EV-BK-001", "تأیید قواعد؛ L4 باز"],
    ["D-01", "confirm", "مدارک: کارت ملی + جواز/اجاره؛ رتبه A/B یا ضامن", "EV-BK-001", "تأیید"],
    ["D-02", "insufficient", "مهلت پیشنهاد/نگهداری موجودی باروک‌ترید", "", "شاهد ناکافی"],
    ["D-03", "insufficient", "تخفیف پلکانی پس از کسری", "", "شاهد ناکافی"],
    ["D-04", "refine", "ذی‌نفع چک ضمانت بانک/طرح؛ چک خرید کالا مسیر جدا", "EV-BK-001", "اصلاح"],
    ["D-05", "partial", "مصرف در کشیدن کارت؛ رزرو پیش از کشیدن در متون عمومی مبهم", "EV-BK-009", "شاهد ناکافی برای رزرو دو سفارش همزمان"],
    ["D-06", "partner_only", "مرجوعی هاگ قبل/بعد سررسید", "EV-BK-009", "نیاز قاعده عمومی"],
    ["D-07", "insufficient", "نقش عملیات باروک‌ترید در اختلاف", "", "شاهد ناکافی"],
    ["D-08", "confirm", "پرداخت هنگام تحویل با پوز باروک مسیر اصلی FMCG", "EV-BK-001;EV-BK-011", "تأیید"],
    ["D-09", "conflict", "برگشت سقف پس از بازپرداخت ادعا شده؛ اصل قسط vs کل پرداخت باز", "EV-BK-006;EV-BK-009", "تعارض باز"],
]


def metrics():
    return {
        "announced_partnerships": 2,  # snapp + saye wallet event
        "listed_acceptors_fmcg": 10,
        "documented_capabilities": 8,
        "forms_or_entry_seen": 2,  # app store + app shell + fmcg form CTA
        "operational_steps_documented": len([x for x in ORDER_STEPS if x[7] in ("documented", "documented_partner", "training_video")]),
        "full_deal_L4": 0,
        "evidence": len(EV),
        "relations": len(RELATIONS),
        "conflicts_open": len([c for c in CONFLICTS if c[3] in ("OPEN", "OPEN_wording", "PARTIAL")]),
    }


def write_excel(m):
    wb = Workbook()
    ws = wb.active
    ws.title = "00_Cover"
    for k, v in m.items():
        ws.append([k, v])
    ws.append(["observed_at", OBS])
    ws.append(["note", "L4 full deal chain not observed; public product+partner guides are main evidence"])

    sheet(wb.create_sheet("01_Entities"), ["id", "kind", "name", "legal_or_brand", "role", "source"], ENTITIES, [14, 16, 24, 22, 40, 20])
    sheet(wb.create_sheet("02_Products"), ["id", "name", "owner", "scope", "status", "evidence"], PRODUCTS, [14, 28, 18, 36, 18, 24])
    sheet(wb.create_sheet("03_PartnerRelations"), ["id", "source", "target", "type", "detail", "proof", "evidence", "vs_prior"], RELATIONS, [12, 16, 16, 24, 36, 8, 24, 16])
    sheet(wb.create_sheet("04_Timeline"), ["date", "event", "kind", "evidence"], TIMELINE, [12, 50, 16, 16])
    sheet(wb.create_sheet("05_Claims"), ["id", "product_or_rel", "audience", "topic", "claim", "as_of", "evidence", "quality", "notes"], CLAIMS, [10, 16, 14, 16, 50, 12, 24, 8, 30])
    sheet(wb.create_sheet("06_Evidence"), ["id", "url", "observed", "title", "snippet", "type", "quality"], EV, [12, 40, 12, 28, 50, 18, 8])
    sheet(wb.create_sheet("07_OrderSteps"), ["id", "scenario", "stage", "role", "path", "action", "response", "status", "evidence", "anon_id", "stop"], ORDER_STEPS, [10, 14, 14, 14, 22, 40, 30, 14, 20, 14, 28])
    sheet(wb.create_sheet("08_PaymentFields"), ["id", "method", "instrument", "payee_or_debit", "certainty_time", "evidence", "notes"], PAY, [18, 14, 22, 30, 30, 14, 36])
    sheet(wb.create_sheet("09_Conflicts"), ["id", "topic", "positions", "status", "ruling_or_next", "evidence"], CONFLICTS, [10, 22, 50, 16, 40, 24])
    sheet(wb.create_sheet("10_CrawlLog"), ["id", "round", "question", "query", "url", "result", "detail", "evidence", "next"], CRAWL, [10, 8, 22, 22, 36, 14, 36, 14, 20])
    sheet(wb.create_sheet("11_Gaps"), ["id", "pri", "topic", "detail", "owner", "artifact_needed", "impact"], GAPS, [10, 6, 28, 40, 18, 24, 14])
    sheet(wb.create_sheet("12_ScenarioDelta"), ["id", "verdict", "proposed_text", "evidence", "ruling"], SCENARIO_DELTA, [12, 14, 55, 24, 28])

    out = ART / "Barook_Kasbno_Deep_Research_Directory.xlsx"
    wb.save(out)
    wb.save(ROOT / out.name)
    return out


def write_scenario_md():
    p = ART / "Barook_Trade_Scenario_v0.3_Research.md"
    lines = [
        "# سناریوی پژوهشی باروک‌ترید نسخهٔ ۰٫۳",
        "",
        f"**تاریخ:** {OBS} · بر اساس شواهد عمومی باروک/کسب‌نو (نه مشاهده تراکنش L4)",
        "",
        "## حکم کلی",
        "مسیر واقعی سوپرمارکت برای موجودی+اعتبار امروز بیشتر «پذیرنده پخش طرف‌قرارداد + باروک‌کارت + پوز موزع + تسویه بانک تجارت» است تا بازارگاه چندپخشیِ کاملِ باروک‌ترید. نسخهٔ ۰٫۲ به‌عنوان فرضیه طراحی حفظ می‌شود؛ موارد زیر حکم پژوهشی‌اند.",
        "",
        "| شناسه | حکم | متن پیشنهادی پژوهشی | شاهد |",
        "| --- | --- | --- | --- |",
    ]
    for r in SCENARIO_DELTA:
        lines.append(f"| {r[0]} | {r[1]} | {r[2]} | {r[3]} |")
    lines += [
        "",
        "## گزینه‌های اتصال باروک‌ترید",
        "1. **ارجاع:** فروشگاه در باروک اعتبار می‌گیرد؛ سفارش در سیستم پخش/شریک می‌ماند؛ پرداخت با کارت روی پوز. پیش‌شرط: لیست پذیرنده و آموزش پوز (موجود).",
        "2. **دستی معتبر:** آپلود/ثبت فاکتور تأمین‌کننده در پنل باروک (آموزش تلگرام) بدون API سفارش.",
        "3. **اتصال مستقیم:** API استعلام سقف/رزرو/مصرف — مستند عمومی کافی نیست (Gap).",
        "",
        "## تعارض‌های باز برای صاحب تصمیم",
        "- CF-02 بازاریابی بدون چک vs چک ۱۲۵٪ رسمی",
        "- CF-04 مبنای شارژ مجدد اعتبار",
        "- CF-05 تعرفه کارمزد باروک",
        "- CF-07 عمق عملیاتی اسنپ‌مارکت پرو",
        "",
    ]
    text = "\n".join(lines)
    p.write_text(text, encoding="utf-8")
    (ROOT / p.name).write_text(text, encoding="utf-8")
    return p


def write_html_pdf(m):
    html_path = ART / "Barook_Kasbno_Deep_Research_FA.html"
    rel_rows = "".join(
        f"<tr><td>{a}</td><td>{b}</td><td>{c}</td><td>{d}</td><td>{e}</td><td>{f}</td><td class='ltr'>{g}</td></tr>"
        for a, b, c, d, e, f, g, h in RELATIONS
    )
    step_rows = "".join(
        f"<tr><td>{a}</td><td>{b}</td><td>{c}</td><td>{e}</td><td>{f}</td><td>{h}</td><td class='ltr'>{i}</td></tr>"
        for a, b, c, d, e, f, g, h, i, j, k in ORDER_STEPS
    )
    cf_rows = "".join(
        f"<tr><td>{a}</td><td>{b}</td><td>{c}</td><td>{d}</td><td>{e}</td></tr>" for a, b, c, d, e, f in CONFLICTS
    )
    html = f"""<!DOCTYPE html>
<html lang="fa" dir="rtl"><head><meta charset="utf-8"/>
<title>پژوهش عمیق باروک / کسب‌نو</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;700&display=swap');
body{{font-family:Vazirmatn,Tahoma,sans-serif;margin:24px;line-height:1.75;color:#122}}
h1,h2,h3{{color:#0B3D5C}} table{{border-collapse:collapse;width:100%;font-size:13px;margin:12px 0}}
th,td{{border:1px solid #ccc;padding:6px;vertical-align:top}} th{{background:#0B3D5C;color:#fff}}
.ltr{{direction:ltr;unicode-bidi:isolate;text-align:left}}
.box{{background:#f4f8fb;padding:12px;border-radius:8px}}
.warn{{background:#fff7e6;border:1px solid #f0d78c;padding:10px}}
.page{{page-break-before:always}}
a{{color:#0B5CAB}}
</style></head><body>
<h1>پژوهش عمیق باروک، کسب‌نو و خرید اعتباری موجودی سوپرمارکت</h1>
<p>تاریخ مشاهده: {OBS}. تمرکز: تأمین موجودی از پخش با اعتبار بانک تجارت از مسیر باروک.</p>
<div class="box">
<p>همکاری اعلام‌شده شاخص: {m['announced_partnerships']} | پذیرنده نام‌برده در FMCG: {m['listed_acceptors_fmcg']} | گام عملیاتی مستند: {m['operational_steps_documented']} | معامله کامل L4: {m['full_deal_L4']}</p>
<p>شواهد: {m['evidence']} | روابط: {m['relations']} | تعارض باز/جزئی: {m['conflicts_open']}</p>
</div>
<div class="warn"><strong>حد پژوهش:</strong> صفحهٔ رسمی و راهنمای شریک و آموزش تلگرام باز شد. حساب پژوهشی و سند بی‌نام یک سفارش از سقف تا برگشت سقف در اختیار نبود؛ بنابراین L4=۰.</div>

<h2>۱. جمع‌بندی مدیریتی</h2>
<p>باروک (بامداد رویان کیان) بازارگاه اعتباری است و برای سوپرمارکت طرح <strong>کسب‌نو بانک تجارت</strong> را اجرا می‌کند: ثبت‌نام و مدارک در باروک، شعبه تجارت، صدور باروک‌کارت، خرید فقط از پخش طرف‌قرارداد، کشیدن کارت روی پوز باروک موزع، کسر خودکار اصل+کارمزد ۲٫۲۵٪ از حساب جاری تجارت، و برگشت مانده پس از بازپرداخت. این نزدیک‌ترین الگوی عمومی «موجودی+اعتبار همان خرید» برای FMCG است.</p>
<p>مدل مجزا: کیف اختصاصی سایه‌سمن (ریسک پخش، بدون اعتبارسنجی اولیه). اعلام اسنپ‌مارکت هنوز در سطح همکاری است نه UI تسویه مشاهده‌شده روی پرو.</p>

<h2>۲. نقشه بازیگران و طرح‌ها</h2>
<ul>
<li><strong>بانک تجارت / اکوتام:</strong> کسب‌نو برای FMCG؛ با تأمین‌نو/صنف‌نو/کشت‌نو قاطی نشود.</li>
<li><strong>باروک:</strong> تجمیع، کارت، پوز، پنل تأمین‌کننده، گزارش.</li>
<li><strong>پذیرندگان FMCG نام‌برده:</strong> عظیمی، گلستان، طبیعت، به پخش، هایمارت، السا، خوانسار، گل‌پخش اول، سایه‌سمن، گلرنگ پخش.</li>
<li><strong>هاگ:</strong> عمیق‌ترین راهنمای عملیاتی کشیدن کارت و ۳۰/۶۰/۹۰.</li>
</ul>

<div class="page"></div>
<h2>۳. مسیر کارت‌خوان (مسیر غالب مشاهده‌شده)</h2>
<ol>
<li>اعتبار و کارت از مسیر باروک+شعبه.</li>
<li>سفارش با ویزیتور یا آنلاین پذیرنده.</li>
<li>سررسید روی فاکتور قبل از کشیدن کارت.</li>
<li>موزع پوز باروک؛ لیست پرداخت‌نشده یا کشیدن مستقیم.</li>
<li>بانک در سررسید حساب را بدهکار می‌کند؛ پخش طبق ادعا از بانک وصول می‌کند.</li>
</ol>
<p>مسیر آنلاین خالص داخل app.barook.ir بدون لاگین دیده نشد. آموزش «خرید/فروش اعتباری بدون حضور» در کانال آمده ولی جزئیات معامله استخراج کامل نشده است.</p>

<h2>۴. روابط</h2>
<table><thead><tr><th>ID</th><th>مبدأ</th><th>مقصد</th><th>نوع</th><th>شرح</th><th>سطح</th><th>شاهد</th></tr></thead><tbody>{rel_rows}</tbody></table>

<div class="page"></div>
<h2>۵. گام‌های سفارش/اعتبار</h2>
<table><thead><tr><th>ID</th><th>سناریو</th><th>مرحله</th><th>مسیر</th><th>اقدام</th><th>وضعیت</th><th>شاهد</th></tr></thead><tbody>{step_rows}</tbody></table>

<div class="page"></div>
<h2>۶. تعارض‌ها</h2>
<table><thead><tr><th>ID</th><th>موضوع</th><th>مواضع</th><th>وضعیت</th><th>حکم/بعدی</th></tr></thead><tbody>{cf_rows}</tbody></table>

<h2>۷. درخواست از متولیان</h2>
<ul>
<li><strong>باروک:</strong> نمونه فاکتور/رسید بی‌نام یک خرید؛ جدول کارمزد باروک؛ ترانسکریپت آموزش پوز.</li>
<li><strong>بانک تجارت (کسب‌نو):</strong> بند شارژ مجدد (اصل یا اصل+کارمزد)؛ صفحه رسمی پایدار PDF.</li>
<li><strong>اسنپ:</strong> راهنمای فعال‌سازی اعتبار روی مارکت/پرو.</li>
<li><strong>سایه‌سمن:</strong> قرارداد ریسک کیف ۵۰م و مرز با کسب‌نو بانک.</li>
</ul>

<h2>۸. فایل‌ها</h2>
<p class="ltr">Excel: artifacts/Barook_Kasbno_Deep_Research_Directory.xlsx</p>
<p class="ltr">Scenario 0.3: artifacts/Barook_Trade_Scenario_v0.3_Research.md</p>
<p class="ltr">PDF: artifacts/Barook_Kasbno_Deep_Research_FA.pdf</p>
<p>پایان — هر تغییر سناریو باید به EV همان گام برسد.</p>
</body></html>"""
    html_path.write_text(html, encoding="utf-8")
    (ROOT / html_path.name).write_text(html, encoding="utf-8")
    pdf_path = ART / "Barook_Kasbno_Deep_Research_FA.pdf"
    printed = False
    for browser in [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]:
        if Path(browser).exists():
            try:
                subprocess.run(
                    [browser, "--headless", "--disable-gpu", f"--print-to-pdf={pdf_path}", html_path.resolve().as_uri()],
                    check=True,
                    timeout=120,
                )
                printed = pdf_path.exists() and pdf_path.stat().st_size > 1000
                if printed:
                    break
            except Exception as ex:
                print("pdf fail", ex)
    if printed:
        import shutil

        shutil.copy(pdf_path, ROOT / pdf_path.name)
    return html_path, pdf_path if printed else None


def main():
    m = metrics()
    xlsx = write_excel(m)
    md = write_scenario_md()
    html, pdf = write_html_pdf(m)
    meta = {
        "observed_at": OBS,
        "metrics": m,
        "xlsx": str(xlsx),
        "scenario_md": str(md),
        "html": str(html),
        "pdf": str(pdf) if pdf else None,
    }
    (ROOT / "build_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (ART / "Barook_Kasbno_Deep_Research_build_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    # dump core json
    (ROOT / "evidence.json").write_text(json.dumps(EV, ensure_ascii=False, indent=2), encoding="utf-8")
    (ROOT / "relations.json").write_text(json.dumps(RELATIONS, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
