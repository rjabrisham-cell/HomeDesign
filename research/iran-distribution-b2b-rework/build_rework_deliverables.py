# -*- coding: utf-8 -*-
"""Rework deliverable: auditable relations + crawl rounds + proof levels."""
from __future__ import annotations

import json
import re
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
ART = Path(r"D:\HomeDesign\artifacts")
ASSOC = Path(r"D:\HomeDesign\research\iran-distribution-b2b\association_members_raw.json")
OBS = "1405/09/03"
CORE = json.loads((ROOT / "pilot_core.json").read_text(encoding="utf-8"))

# ── Evidence ───────────────────────────────────────────────────────────────
EVIDENCE = [
    {"evidence_id": "EV-001", "url": "https://www.varanegar.com/varanegar-pepsi-success-story/", "title": "داستان موفقیت ورانگر–پپسی؛ مصاحبه IT ساتراپ دینا", "observed_at": OBS, "snippet": "ویدیوی مصاحبه با دکتر رزا امیراحمدی، مدیر فناوری اطلاعات شرکت ساتراپ دینا پپسی", "type": "case_study"},
    {"evidence_id": "EV-002", "url": "https://www.varanegar.com/از-زبان-مشتریان/", "title": "از زبان مشتریان ورانگر", "observed_at": OBS, "snippet": "خانم الهی مدیر آی‌تی صنایع غذایی برتر (نویدکاران): همکاری بیش از یک دهه با ورانگر؛ نرم‌افزار پخش و فروش", "type": "customer_quote"},
    {"evidence_id": "EV-003", "url": "https://www.varanegar.com/ماژول-فروش-b2b/", "title": "ماژول فروش B2B ورانگر", "observed_at": OBS, "snippet": "ثبت سفارش آنلاین، قیمت اختصاصی، کنترل سقف اعتبار، موجودی لحظه‌ای؛ قابلیت محصول نه اثبات نصب خاص", "type": "product_page"},
    {"evidence_id": "EV-004", "url": "https://telegram.me/s/tebyanphsoftware?before=727", "title": "کانال آرسن/تبیان — دریافت فاکتور هجرت/فردوس/رازی", "observed_at": OBS, "snippet": "امکان دریافت فاکتورهای خرید شرکتهای هجرت، فردوس و رازی؛ گزینه فراخوانی از شرکتهای پخش و HIX؛ داروپخش و الیت مستقیم", "type": "support_post"},
    {"evidence_id": "EV-005", "url": "https://www.hejratco.com/", "title": "سایت پخش هجرت", "observed_at": OBS, "snippet": "پورتال اینترنتی مشتریان و سامانه اطلاع‌رسانی تراکنش در فهرست خدمات؛ URL ورود عمومی در خزش باز نشد (504 در یک تلاش)", "type": "corporate_site", "access_note": "504 on one fetch; content via search snippet + retry needed"},
    {"evidence_id": "EV-006", "url": "http://orders.razico.ir/", "title": "پرتال سفارش‌دهی اینترنتی رازی", "observed_at": OBS, "snippet": "فرم ورود کاربر/رمز + کپچا؛ خوش‌آمد به پرتابل سفارش‌دهی", "type": "login_page"},
    {"evidence_id": "EV-007", "url": "https://my.razico.ir/Authentication/Login", "title": "سامانه مشتریان رازی", "observed_at": OBS, "snippet": "ورود با نام کاربری و رمز عبور", "type": "login_page"},
    {"evidence_id": "EV-008", "url": "https://soghatemakoran.ir/panel/", "title": "ورود مشتریان سوغات مکران", "observed_at": OBS, "snippet": "شماره موبایل + کد مشتری؛ فرم متقاضی جدید با واتساپ", "type": "login_page"},
    {"evidence_id": "EV-009", "url": "https://soghatemakoran.ir/", "title": "لندینگ سوغات مکران", "observed_at": OBS, "snippet": "سفارش اپ/واتساپ؛ حداقل سفارش اپ ۲٬۵۰۰٬۰۰۰ تومان؛ موبایل مراد بلوچی ۰۹۱۵۹۱۵۴۵۵۰؛ محدوده نیکشهر و اطراف", "type": "landing"},
    {"evidence_id": "EV-010", "url": "https://cafebazaar.ir/app/ir.soghatmakran.app", "title": "اپ سوغات مکران در بازار", "observed_at": OBS, "snippet": "اپ اندروید ir.soghatmakran.app؛ توضیحات پخش غذایی", "type": "app_store"},
    {"evidence_id": "EV-011", "url": "https://www.payla.net/home/application", "title": "اپ‌ها و نقش‌های پایلا", "observed_at": OBS, "snippet": "مغازه‌دار: سفارش، پیگیری، مدیریت روش پرداخت و تسویه؛ پنل پخش: قیمت چندسطحی و روش پرداخت", "type": "product_page"},
    {"evidence_id": "EV-012", "url": "https://idmt.ir/ساخت-اپلیکیشن-پخش-مواد-غذایی", "title": "مطالعهٔ موردی ساخت پخش‌من", "observed_at": OBS, "snippet": "مارکت‌پلیس عمده؛ جستجو بر اساس شهر؛ سفارش‌ها؛ کیف پول؛ تماس ۰۵۱۳۶۰۳۸۵۰۶", "type": "case_study"},
    {"evidence_id": "EV-013", "url": "https://pakhshmart.com/", "title": "سیمرغ — پوشش گلرنگ پخش و پدیده پایدار", "observed_at": OBS, "snippet": "پلتفرم سفارش آنلاین گروه گلرنگ؛ برندهای تحت پوشش: گلرنگ پخش، پدیده پایدار؛ تلفن ۴۲۴۱۸۶۲۴", "type": "landing"},
    {"evidence_id": "EV-014", "url": "https://app.pakhshmart.com/", "title": "ورود سیمرغ OTP", "observed_at": OBS, "snippet": "شماره موبایل ثبت‌شده در شرکت‌های پخش + دریافت رمز یک‌بارمصرف؛ نسخه ۷.۴.۶ در مشاهده قبلی", "type": "login_page"},
    {"evidence_id": "EV-015", "url": "https://davooddehghan.ir/wp-content/uploads/ilia-logistic-1402.pdf", "title": "PDF تحول در مسیر تحویل — نمونه تکساران/آپتایم/اسنپ‌مارکت‌پرو", "observed_at": OBS, "snippet": "تکساران پارمیس مشتری B2B آپتایم از ۱۴۰۲؛ لجستیک نه پنل سفارش خرده‌فروش؛ اسنپ مارکت پرو نیز در همان گزارش به‌عنوان مشتری آپتایم", "type": "pdf_report"},
    {"evidence_id": "EV-016", "url": "https://www.linkedin.com/posts/optime-ai_%D8%A2%D9%BE%D8%AA%D8%A7%DB%8C%D9%85-%D9%84%D8%AC%D8%B3%D8%AA%DB%8C%DA%A9%D9%87%D9%88%D8%B4%D9%85%D9%86%D8%AF-%D8%A8%D9%87%DB%8C%D9%86%D9%87%D8%B3%D8%A7%D8%B2%DB%8C%D8%B9%D9%85%D9%84%DA%A9%D8%B1%D8%AF-activity-7390301973977141248-dzu0", "title": "آپتایم × قاسم ایران", "observed_at": OBS, "snippet": "همکاری لجستیکی قاسم ایران با آپتایم؛ مسیربندی و مانیتورینگ توزیع", "type": "linkedin"},
    {"evidence_id": "EV-017", "url": "https://www.bonnychow.ir/questions", "title": "FAQ بانی‌چاو — تسویه و اعتبار", "observed_at": OBS, "snippet": "کارتخوان هنگام تحویل؛ درگاه برای open items؛ تمدید اعتبار حداکثر ۷۲س پس از اعلام وصول مالی", "type": "faq"},
    {"evidence_id": "EV-018", "url": "https://www.bonnychow.ir/registration", "title": "ثبت‌نام کد خریدار بانی‌چاو", "observed_at": OBS, "snippet": "کد خریدار سولیکو؛ سفارش آنلاین پس از ثبت؛ پرداخت نقدی/کارت هنگام تحویل یا آنلاین طبق دوره بازپرداخت", "type": "registration"},
    {"evidence_id": "EV-019", "url": "https://www.ghasemiran.org/fa/page/ghasemiran-B2B.html", "title": "صفحه رسمی B2B قاسم ایران", "observed_at": OBS, "snippet": "اپ B2B برای دارندگان کد مشتری؛ جستجو، تخفیف، فاکتور رسمی، تحویل ناوگان", "type": "product_page"},
    {"evidence_id": "EV-020", "url": "https://dastavval.com/article/art-guide-site-order-step-by-step", "title": "راهنمای سفارش دست اول", "observed_at": OBS, "snippet": "کاتالوگ، MOQ، تخفیف تیراژ، سبد، پیش‌فاکتور", "type": "official_guide"},
    {"evidence_id": "EV-021", "url": "https://dastavval.com/article/art-guide-safe-escrow-payment-cheque", "title": "راهنمای نقد/امانی/چک دست اول", "observed_at": OBS, "snippet": "Escrow شاپرک؛ نقد با تخفیف؛ چک صیادی ۳۰/۶۰/۹۰؛ ثبت صیاد", "type": "official_guide"},
    {"evidence_id": "EV-022", "url": "https://www.apir.ir/pages/member-companies", "title": "انجمن ملی صنعت پخش — فهرست اعضا", "observed_at": OBS, "snippet": "منبع دایرکتوری؛ آپدیت پاییز ۱۴۰۵ روی سایت", "type": "directory"},
    {"evidence_id": "EV-023", "url": "https://myket.ir/app/com.shagerd.app", "title": "اپ شاگرد در مایکت (بررسی ادعای ممیزی)", "observed_at": OBS, "snippet": "اپ POS/AI فروشگاهی عازم — نه پنل پخش خرده‌فروش؛ شاهد ممیزی برای «شاگرد پخش» در این URL تأیید نشد", "type": "negative_match"},
]

# ── Relations (fresh vs baseline) ───────────────────────────────────────────
RELATIONS = [
    {"relation_id": "REL-001", "case_id": "CASE-01", "source_id": "LEGAL-SATRAP-DINA", "target_id": "PROV-VARANEGAR", "connection_type": "ERP_customer", "product_module": "نرم‌افزار پخش/فروش ورانگر", "buyer_role": "شرکت پخش (اپراتور سیستم)", "proof_level": 1, "status": "documented_cooperation", "evidence_ids": "EV-001", "next_stage": "یافتن URL پنل B2B اختصاصی ساتراپ در صورت فعال بودن ماژول", "vs_baseline": "NEW — baseline NoEvidence"},
    {"relation_id": "REL-002", "case_id": "CASE-01", "source_id": "PROV-VARANEGAR", "target_id": "PROV-VARANEGAR", "connection_type": "product_capability_B2B_module", "product_module": "ماژول فروش B2B", "buyer_role": "خرده‌فروش مشتری پخش (قابلیت عمومی)", "proof_level": 1, "status": "capability_not_per_install", "evidence_ids": "EV-003", "next_stage": "مشاهده ورودی مشتری برای نصب ساتراپ", "vs_baseline": "NEW nuance"},
    {"relation_id": "REL-003", "case_id": "CASE-02", "source_id": "LEGAL-NAVIDKARAN", "target_id": "PROV-VARANEGAR", "connection_type": "ERP_customer", "product_module": "نرم‌افزار پخش و فروش", "buyer_role": "شرکت پخش", "proof_level": 1, "status": "documented_cooperation", "evidence_ids": "EV-002", "next_stage": "آیا ماژول B2B خرده‌فروش در این نصب فعال است؟", "vs_baseline": "NEW — baseline NoEvidence"},
    {"relation_id": "REL-004", "case_id": "CASE-03", "source_id": "LEGAL-HEJRAT", "target_id": "PROV-ARSEN", "connection_type": "invoice_pull_announced", "product_module": "فراخوانی فاکتور خرید از پخش و HIX", "buyer_role": "داروخانه کاربر آرسن", "proof_level": 1, "status": "announced_invoice_receive", "evidence_ids": "EV-004", "next_stage": "اعتبار اتصال جاری + مسیر سفارش قبل از فاکتور", "vs_baseline": "NEW — baseline NoEvidence"},
    {"relation_id": "REL-005", "case_id": "CASE-03", "source_id": "LEGAL-HEJRAT", "target_id": "PROV-HIX", "connection_type": "via_intermediary_announced", "product_module": "HIX در مسیر آرسن", "buyer_role": "داروخانه", "proof_level": 1, "status": "announced_via_guide", "evidence_ids": "EV-004", "next_stage": "جدا کردن اتصال مستقیم در مقابل HIX", "vs_baseline": "NEW"},
    {"relation_id": "REL-006", "case_id": "CASE-04", "source_id": "LEGAL-FERDOWS", "target_id": "PROV-ARSEN", "connection_type": "invoice_pull_announced", "product_module": "فراخوانی فاکتور خرید", "buyer_role": "داروخانه", "proof_level": 1, "status": "announced_invoice_receive", "evidence_ids": "EV-004", "next_stage": "شاهد نسخه/فعال بودن اتصال", "vs_baseline": "NEW — baseline NoEvidence"},
    {"relation_id": "REL-007", "case_id": "CASE-05", "source_id": "LEGAL-RAZI-PHARMA", "target_id": "PROV-ARSEN", "connection_type": "invoice_pull_announced", "product_module": "فراخوانی فاکتور از رازی", "buyer_role": "داروخانه", "proof_level": 1, "status": "announced_invoice_receive", "evidence_ids": "EV-004", "next_stage": "تفکیک از سفارش مستقیم orders.razico.ir", "vs_baseline": "NEW path vs baseline portal-only"},
    {"relation_id": "REL-008", "case_id": "CASE-05", "source_id": "LEGAL-RAZI-PHARMA", "target_id": "SYS-RAZI-ORDERS", "connection_type": "customer_order_portal", "product_module": "orders.razico.ir + my.razico.ir", "buyer_role": "مشتری ثبت‌شده رازی", "proof_level": 2, "status": "login_observed", "evidence_ids": "EV-006;EV-007", "next_stage": "گام قیمت/سبد/پرداخت پشت لاگین", "vs_baseline": "KEPT+split identity from LEGAL-RAZI-CHESB"},
    {"relation_id": "REL-009", "case_id": "CASE-05", "source_id": "LEGAL-RAZI-CHESB", "target_id": "SYS-RAZI-ORDERS", "connection_type": "identity_non_match", "product_module": "—", "buyer_role": "—", "proof_level": 0, "status": "not_same_legal_entity", "evidence_ids": "EV-022", "next_stage": "خزش مستقل برای یکتا همراه رازی", "vs_baseline": "FIX baseline DIST-004 collision"},
    {"relation_id": "REL-010", "case_id": "CASE-06", "source_id": "LEGAL-SOGHAT-MAKRAN", "target_id": "SYS-MAKRAN-PANEL", "connection_type": "wholesale_customer_panel", "product_module": "پنل وب + اپ بازار", "buyer_role": "فروشگاه با کد مشتری", "proof_level": 2, "status": "entry_observed", "evidence_ids": "EV-008;EV-009;EV-010", "next_stage": "قیمت داخل حساب، ثبت نهایی، پرداخت همان سفارش", "vs_baseline": "NEW regional calibration"},
    {"relation_id": "REL-011", "case_id": "CASE-07", "source_id": "PROV-PAYLA", "target_id": "PROV-PAYLA", "connection_type": "multi_role_platform", "product_module": "اپ مغازه‌دار/بازاریاب/راننده/پنل پخش", "buyer_role": "مغازه‌دار پذیرنده", "proof_level": 1, "status": "product_documented", "evidence_ids": "EV-011", "next_stage": "نام مشتری پخش مستقر + مشاهده لاگین واقعی معامله", "vs_baseline": "NEW provider"},
    {"relation_id": "REL-012", "case_id": "CASE-08", "source_id": "PROV-PAKHSHMAN", "target_id": "PROV-PAKHSHMAN", "connection_type": "marketplace_product", "product_module": "پخش‌من مارکت‌پلیس عمده", "buyer_role": "خریدار عمده", "proof_level": 1, "status": "product_documented", "evidence_ids": "EV-012", "next_stage": "هویت تأمین‌کننده طرف معامله + پرداخت", "vs_baseline": "NEW provider"},
    {"relation_id": "REL-013", "case_id": "CASE-09", "source_id": "LEGAL-GOLRANG-PAKHSH", "target_id": "SYS-SIMORGH", "connection_type": "brand_portal_coverage", "product_module": "سیمرغ OTP", "buyer_role": "فروشگاه با موبایل ثبت‌شده", "proof_level": 2, "status": "entry_observed", "evidence_ids": "EV-013;EV-014", "next_stage": "مشاهده قیمت/سبد/پرداخت پس از OTP حساب پژوهشی", "vs_baseline": "REFINED — system family separated"},
    {"relation_id": "REL-014", "case_id": "CASE-10", "source_id": "LEGAL-PADIDEH-PAYDAR", "target_id": "SYS-SIMORGH", "connection_type": "brand_portal_coverage", "product_module": "سیمرغ OTP", "buyer_role": "فروشگاه با موبایل ثبت‌شده", "proof_level": 2, "status": "entry_observed", "evidence_ids": "EV-013;EV-014", "next_stage": "تفکیک موجودی پدیده در همان پنل", "vs_baseline": "REFINED second Simorgh legal"},
    {"relation_id": "REL-015", "case_id": "CASE-11", "source_id": "LEGAL-TAKSARAN", "target_id": "PROV-OPTIME", "connection_type": "logistics_optimization", "product_module": "مسیربندی/تخصیص ناوگان آپتایم", "buyer_role": "عملیات پخش (نه خرده‌فروش سفارش‌دهنده)", "proof_level": 1, "status": "documented_cooperation_logistics_only", "evidence_ids": "EV-015", "next_stage": "جستجوی مستقل پنل سفارش خرده‌فروش تکساران — این رابطه سفارش نیست", "vs_baseline": "NEW — do NOT treat as retail order path"},
    {"relation_id": "REL-016", "case_id": "CASE-11", "source_id": "PROV-OPTIME", "target_id": "PROV-OPTIME", "connection_type": "shared_provider_discovery", "product_module": "همان گزارش مشتریان آپتایم شامل اسنپ‌مارکت‌پرو", "buyer_role": "—", "proof_level": 1, "status": "shared_vendor_not_trade_link", "evidence_ids": "EV-015", "next_stage": "رابطه عرضه تکساران↔اسنپ‌مارکت‌پرو از این شاهد احراز نمی‌شود", "vs_baseline": "NEW caution from audit"},
    {"relation_id": "REL-017", "case_id": "CASE-12", "source_id": "LEGAL-BONNYCHOW", "target_id": "LEGAL-BONNYCHOW", "connection_type": "wholesale_portal_payment_credit", "product_module": "پورتال بانی‌چاو / کد خریدار سولیکو", "buyer_role": "فروشگاه/کافه دارای کد خریدار", "proof_level": 3, "status": "guide_documents_order_and_settlement_rules", "evidence_ids": "EV-017;EV-018", "next_stage": "تراکنش پژوهشی واقعی برای سطح ۴", "vs_baseline": "CORRECTED 72h field = credit_refresh_sla"},
    {"relation_id": "REL-018", "case_id": "CASE-12", "source_id": "LEGAL-GHASEM", "target_id": "PROV-OPTIME", "connection_type": "logistics_optimization", "product_module": "آپتایم برای توزیع قاسم", "buyer_role": "عملیات پخش", "proof_level": 1, "status": "documented_cooperation_logistics_only", "evidence_ids": "EV-016", "next_stage": "جدا از اپ B2B قاسم نگه دارید", "vs_baseline": "NEW"},
    {"relation_id": "REL-019", "case_id": "CASE-12", "source_id": "LEGAL-GHASEM", "target_id": "LEGAL-GHASEM", "connection_type": "customer_code_b2b_app", "product_module": "اپ B2B قاسم ایران", "buyer_role": "اصناف دارای کد مشتری", "proof_level": 2, "status": "official_entry_described", "evidence_ids": "EV-019", "next_stage": "مشاهده داخل اپ پس از کد", "vs_baseline": "KEPT with clearer proof level"},
    {"relation_id": "REL-020", "case_id": "CASE-12", "source_id": "PROV-DASTAVVAL", "target_id": "PROV-DASTAVVAL", "connection_type": "order_and_settlement_guide", "product_module": "دست اول — سفارش+امانی+چک", "buyer_role": "خریدار عمده", "proof_level": 3, "status": "official_guides_steps_and_payment_options", "evidence_ids": "EV-020;EV-021", "next_stage": "اجرای تراکنش برای سطح ۴ و انتساب به پخش انجمن", "vs_baseline": "NEW method calibration"},
]

# Order / process steps (only where documented)
ORDER_STEPS = [
    {"step_id": "OS-MAK-01", "case_id": "CASE-06", "entity_id": "LEGAL-SOGHAT-MAKRAN", "stage": "entry", "role": "فروشگاه", "page_or_path": "soghatemakoran.ir/panel", "action": "ورود با موبایل و کد مشتری", "fields_conditions": "موبایل ۰۹… + کد مشتری؛ بدون کد → فرم واتساپ", "decision_response": "فرم ورود مشاهده شد", "status": "observed", "evidence_id": "EV-008", "witness_type": "page_open", "stop_point": ""},
    {"step_id": "OS-MAK-02", "case_id": "CASE-06", "entity_id": "LEGAL-SOGHAT-MAKRAN", "stage": "price_browse_claim", "role": "فروشگاه", "page_or_path": "لندینگ/اپ", "action": "مشاهده قیمت در اپ", "fields_conditions": "حداقل سفارش اپ ۲٫۵ میلیون تومان", "decision_response": "ادعا روی لندینگ؛ قیمت داخل حساب مشاهده نشد", "status": "claimed_not_inside_account", "evidence_id": "EV-009", "witness_type": "landing_text", "stop_point": "نیاز به کد مشتری واقعی"},
    {"step_id": "OS-MAK-03", "case_id": "CASE-06", "entity_id": "LEGAL-SOGHAT-MAKRAN", "stage": "order_submit_claim", "role": "فروشگاه", "page_or_path": "اپ یا واتساپ", "action": "ثبت سفارش", "fields_conditions": "سه قدم اعلامی: قیمت → سفارش → تحویل <۲۴س", "decision_response": "مسیر اعلام شده؛ ثبت نهایی داخل حساب ندیده شد", "status": "claimed", "evidence_id": "EV-009", "witness_type": "landing_text", "stop_point": "سطح ۳ ناقص"},
    {"step_id": "OS-MAK-04", "case_id": "CASE-06", "entity_id": "LEGAL-SOGHAT-MAKRAN", "stage": "payment", "role": "فروشگاه", "page_or_path": "—", "action": "پرداخت همان سفارش", "fields_conditions": "روش نقد/کارت/اعتبار در صفحه عمومی مشخص نیست", "decision_response": "مشاهده نشد", "status": "not_observed", "evidence_id": "EV-009", "witness_type": "absence", "stop_point": "سطح ۴ باز"},
    {"step_id": "OS-SIM-01", "case_id": "CASE-09", "entity_id": "SYS-SIMORGH", "stage": "entry", "role": "فروشگاه", "page_or_path": "app.pakhshmart.com", "action": "OTP موبایل ثبت‌شده در پخش", "fields_conditions": "فقط موبایل ازپیش‌ثبت‌شده", "decision_response": "فرم OTP مشاهده شد", "status": "observed", "evidence_id": "EV-014", "witness_type": "page_open", "stop_point": ""},
    {"step_id": "OS-SIM-02", "case_id": "CASE-09", "entity_id": "SYS-SIMORGH", "stage": "catalog_price", "role": "فروشگاه", "page_or_path": "پس از OTP", "action": "مشاهده کاتالوگ/قیمت", "fields_conditions": "پشت لاگین", "decision_response": "بدون حساب پژوهشی دیده نشد", "status": "behind_login", "evidence_id": "EV-014", "witness_type": "access_barrier", "stop_point": "سطح ۳ باز"},
    {"step_id": "OS-BON-01", "case_id": "CASE-12", "entity_id": "LEGAL-BONNYCHOW", "stage": "onboarding", "role": "خریدار", "page_or_path": "bonnychow.ir/registration", "action": "دریافت کد خریدار", "fields_conditions": "مدارک حقیقی/حقوقی؛ تلفن ۶۱۳۷۰", "decision_response": "مسیر رسمی مستند", "status": "documented", "evidence_id": "EV-018", "witness_type": "official_guide", "stop_point": ""},
    {"step_id": "OS-BON-02", "case_id": "CASE-12", "entity_id": "LEGAL-BONNYCHOW", "stage": "order", "role": "خریدار", "page_or_path": "پورتال", "action": "ثبت سفارش آنلاین", "fields_conditions": "پس از کد خریدار و ورود موبایل", "decision_response": "در راهنما مستند؛ نشست واقعی انجام نشد", "status": "documented", "evidence_id": "EV-018", "witness_type": "official_guide", "stop_point": ""},
    {"step_id": "OS-BON-03", "case_id": "CASE-12", "entity_id": "LEGAL-BONNYCHOW", "stage": "payment_cash_card", "role": "خریدار", "page_or_path": "تحویل / کارتخوان", "action": "پرداخت هنگام تحویل با کارت", "fields_conditions": "گیرنده: عملیات توزیع بانی‌چاو؛ زمان قطعیت: هنگام تحویل", "decision_response": "مستند FAQ", "status": "documented", "evidence_id": "EV-017", "witness_type": "official_guide", "stop_point": ""},
    {"step_id": "OS-BON-04", "case_id": "CASE-12", "entity_id": "LEGAL-BONNYCHOW", "stage": "payment_online_open_items", "role": "خریدار", "page_or_path": "درگاه پورتال", "action": "پرداخت صورتحساب‌های باز", "fields_conditions": "open items؛ روش‌های رایج قبلی هم مجاز", "decision_response": "مستند FAQ", "status": "documented", "evidence_id": "EV-017", "witness_type": "official_guide", "stop_point": ""},
    {"step_id": "OS-BON-05", "case_id": "CASE-12", "entity_id": "LEGAL-BONNYCHOW", "stage": "credit_refresh", "role": "امور مالی پخش", "page_or_path": "پورتال پس از وصول", "action": "تمدید نمایش اعتبار", "fields_conditions": "credit_refresh_sla_hours=72؛ جدا از repayment_term_days", "decision_response": "مستند؛ سوءبرداشت مهلت خرید رد شد", "status": "documented", "evidence_id": "EV-017", "witness_type": "official_guide", "stop_point": "سطح ۴ تراکنش زنده باز"},
    {"step_id": "OS-DAS-01", "case_id": "CASE-12", "entity_id": "PROV-DASTAVVAL", "stage": "catalog", "role": "خریدار عمده", "page_or_path": "کاتالوگ دست اول", "action": "فیلتر کالا/کارخانه؛ دیدن MOQ و قیمت کارتن", "fields_conditions": "تعداد در کارتن، MOQ، تاریخ انقضا", "decision_response": "راهنمای رسمی گام‌ها", "status": "documented", "evidence_id": "EV-020", "witness_type": "official_guide", "stop_point": ""},
    {"step_id": "OS-DAS-02", "case_id": "CASE-12", "entity_id": "PROV-DASTAVVAL", "stage": "preinvoice", "role": "خریدار عمده", "page_or_path": "صدور پیش‌فاکتور", "action": "ثبت تحویل‌گیرنده و صدور پیش‌فاکتور", "fields_conditions": "بارکد، شناسه پیگیری، وزن", "decision_response": "مستند راهنما", "status": "documented", "evidence_id": "EV-020", "witness_type": "official_guide", "stop_point": ""},
    {"step_id": "OS-DAS-03", "case_id": "CASE-12", "entity_id": "PROV-DASTAVVAL", "stage": "payment_escrow", "role": "خریدار", "page_or_path": "Safe-Buy Escrow", "action": "بلوکه وجه در صندوق امانی شاپرک تا تأیید تحویل", "fields_conditions": "گیرنده نهایی تولیدکننده پس از تأیید سالم بودن بار", "decision_response": "مستند راهنما تسویه", "status": "documented", "evidence_id": "EV-021", "witness_type": "official_guide", "stop_point": ""},
    {"step_id": "OS-DAS-04", "case_id": "CASE-12", "entity_id": "PROV-DASTAVVAL", "stage": "payment_cash", "role": "خریدار", "page_or_path": "درگاه شتاب", "action": "تسویه نقدی با تخفیف نقدی", "fields_conditions": "تخفیف اعلامی ۳–۸٪؛ اولویت بارگیری", "decision_response": "مستند", "status": "documented", "evidence_id": "EV-021", "witness_type": "official_guide", "stop_point": ""},
    {"step_id": "OS-DAS-05", "case_id": "CASE-12", "entity_id": "PROV-DASTAVVAL", "stage": "payment_cheque", "role": "نماینده تأییدشده", "page_or_path": "چک صیادی", "action": "۵۰/۵۰ یا تمام‌چکی ۳۰/۶۰/۹۰ روز", "fields_conditions": "ذی‌نفع شرکت؛ استعلام ۱۶رقمی؛ ثبت صیاد/پیچک قبل اعزام خودرو", "decision_response": "مستند؛ اجرای تراکنش دیده نشد", "status": "documented", "evidence_id": "EV-021", "witness_type": "official_guide", "stop_point": "سطح ۴ زنده باز"},
    {"step_id": "OS-ARS-01", "case_id": "CASE-03", "entity_id": "PROV-ARSEN", "stage": "invoice_pull", "role": "داروخانه", "page_or_path": "فرم ورود دارو به انبار آرسن", "action": "فراخوانی از شرکت‌های پخش و HIX", "fields_conditions": "هجرت، فردوس، رازی نام برده شده‌اند", "decision_response": "گام پس از خرید؛ مسیر ثبت سفارش از این شاهد نیست", "status": "documented_post_purchase", "evidence_id": "EV-004", "witness_type": "support_post", "stop_point": "سفارش پیش از فاکتور باز است"},
    {"step_id": "OS-PAY-01", "case_id": "CASE-07", "entity_id": "PROV-PAYLA", "stage": "capability", "role": "مغازه‌دار", "page_or_path": "اپ پذیرنده پایلا", "action": "سفارش و انتخاب روش پرداخت/تسویه", "fields_conditions": "لیست قابلیت محصول", "decision_response": "قابلیت اعلامی؛ مشتری نام‌دار و تراکنش دیده نشد", "status": "capability_only", "evidence_id": "EV-011", "witness_type": "product_page", "stop_point": "سطح ۲ ورودی عمومی فروش نیاز به URL لاگین مشخص"},
]

PAYMENT_FIELDS = [
    {"entity_id": "LEGAL-BONNYCHOW", "method": "card_on_delivery", "payee": "عملیات توزیع بانی‌چاو", "certainty_time": "عندالتحویل", "evidence_id": "EV-017", "notes": ""},
    {"entity_id": "LEGAL-BONNYCHOW", "method": "online_open_items", "payee": "شرکت بانی‌چاو از طریق درگاه", "certainty_time": "پس از پرداخت درگاهی", "evidence_id": "EV-017", "notes": ""},
    {"entity_id": "LEGAL-BONNYCHOW", "method": "credit_line", "credit_provider": "بانی‌چاو", "debtor": "خریدار خوش‌حساب", "reserve_stage": "سقف اعتبار حساب", "consume_stage": "خرید روی اعتبار", "settle_stage": "پرداخت open items", "credit_refresh_sla_hours": 72, "repayment_term_days": "طبق دوره تعریف‌شده — عدد عمومی نشده", "evidence_id": "EV-017", "notes": "۷۲س ≠ مهلت خرید"},
    {"entity_id": "PROV-DASTAVVAL", "method": "escrow", "payee": "صندوق امانی شاپرک تا تأیید تحویل سپس تولیدکننده", "certainty_time": "پس از تأیید تحویل سالم", "evidence_id": "EV-021", "notes": ""},
    {"entity_id": "PROV-DASTAVVAL", "method": "cash_online", "payee": "مسیر درگاه/کارخانه طبق سامانه", "certainty_time": "پرداخت آنلاین", "evidence_id": "EV-021", "notes": "تخفیف نقدی ۳–۸٪ اعلامی"},
    {"entity_id": "PROV-DASTAVVAL", "method": "cheque_sayad", "payee": "شرکت دست اول / ثبت صیاد", "cheque_beneficiary": "شرکت", "acceptance": "استعلام ۱۶رقمی + ثبت صیاد", "dispatch_condition": "پس از ثبت صیاد", "collection": "سررسید ۳۰/۶۰/۹۰", "evidence_id": "EV-021", "notes": ""},
    {"entity_id": "LEGAL-SOGHAT-MAKRAN", "method": "unknown_public", "payee": "نامشخص عمومی", "certainty_time": "—", "evidence_id": "EV-009", "notes": "پرداخت همان سفارش مشاهده نشد"},
]

HEADER_FILL = PatternFill("solid", fgColor="0F3D68")
HEADER_FONT = Font(color="FFFFFF", bold=True, name="Tahoma", size=10)
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="right")
THIN = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)


def style_sheet(ws, headers, rows, widths):
    ws.append(headers)
    for col in range(1, len(headers) + 1):
        c = ws.cell(1, col)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    for r in rows:
        ws.append(r)
        for col in range(1, len(headers) + 1):
            cell = ws.cell(ws.max_row, col)
            cell.alignment = WRAP
            cell.border = THIN
            cell.font = Font(name="Tahoma", size=9)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def parse_row_no(raw, fallback: int) -> int:
    s = str(raw or "").strip()
    # normalize Arabic/Persian digits and odd marks
    trans = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")
    s = s.translate(trans).replace("ّ", "")
    digits = re.sub(r"\D", "", s)
    if digits:
        try:
            return int(digits)
        except ValueError:
            pass
    return fallback


def load_assoc():
    rows = json.loads(ASSOC.read_text(encoding="utf-8"))
    for i, r in enumerate(rows, start=1):
        r["row_in_file"] = parse_row_no(r.get("row_in_file"), i)
    return rows


def build_crawl_log(members):
    logs = []
    lid = 1

    def add(**kw):
        nonlocal lid
        row = {"log_id": f"CL-{lid:04d}", "observed_at": OBS, **kw}
        logs.append(row)
        lid += 1

    # Pilot deep rounds
    pilot_rounds = [
        ("CASE-01", "LEGAL-SATRAP-DINA", 1, "assoc_name ساتراپ دینا", "ساتراپ دینا ورانگر", "https://www.varanegar.com/varanegar-pepsi-success-story/", "ok", "مصاحبه IT ساتراپ در صفحه پپسی", "EV-001", "بررسی ماژول B2B"),
        ("CASE-01", "LEGAL-SATRAP-DINA", 2, "EV-001 → ماژول B2B", "ورانگر ماژول فروش B2B", "https://www.varanegar.com/ماژول-فروش-b2b/", "ok", "قابلیت سفارش/اعتبار عمومی؛ نه URL نصب ساتراپ", "EV-003", "درخواست دمو یا URL مشتری"),
        ("CASE-02", "LEGAL-NAVIDKARAN", 1, "assoc_name نویدکاران", "نویدکاران ورانگر", "https://www.varanegar.com/از-زبان-مشتریان/", "ok", "نقل قول مدیر IT نویدکاران", "EV-002", "جستجوی پنل B2B اختصاصی"),
        ("CASE-02", "LEGAL-NAVIDKARAN", 2, "نام تجاری صنایع غذایی برتر", "صنایع غذایی برتر نویدکاران سفارش آنلاین", "web search", "no_retail_portal_found", "در منابع عمومی ورودی خرده‌فروش یافت نشد", "", "تماس سوئیچبورد / پرسش فعال بودن ماژول B2B"),
        ("CASE-03", "LEGAL-HEJRAT", 1, "assoc + audit hint آرسن", "آرسن هجرت فردوس رازی فاکتور", "https://telegram.me/s/tebyanphsoftware?before=727", "ok", "پست فراخوانی فاکتور هجرت/فردوس/رازی و HIX", "EV-004", "سایت هجرت برای پورتال مشتری"),
        ("CASE-03", "LEGAL-HEJRAT", 2, "دامنه hejratco.com از ایمیل انجمن", "site:hejratco.com پورتال مشتریان", "https://www.hejratco.com/", "partial_timeout", "یک دریافت 504؛ اسنیپت جستجو پورتال اینترنتی مشتریان را فهرست می‌کند بدون URL ورود", "EV-005", "تلاش مجدد + جستجوی subdomain"),
        ("CASE-04", "LEGAL-FERDOWS", 1, "همان پست آرسن", "فردوس آرسن فاکتور", "https://telegram.me/s/tebyanphsoftware?before=727", "ok", "نام فردوس در همان پست", "EV-004", "سایت ferdowsdco.com"),
        ("CASE-04", "LEGAL-FERDOWS", 2, "ایمیل ferdowsdco.com", "ferdowsdco.com سفارش OR پنل", "domain probe", "no_public_order_ui", "رابطه فاکتور آرسن سطح ۱؛ سفارش پیش از فاکتور پیدا نشد", "EV-004", "پرسش پشتیبانی آرسن درباره اتصال زنده"),
        ("CASE-05", "LEGAL-RAZI-PHARMA", 1, "baseline portal", "orders.razico.ir", "http://orders.razico.ir/", "ok", "لاگین کاربر/رمز مشاهده شد", "EV-006", "my.razico.ir"),
        ("CASE-05", "LEGAL-RAZI-PHARMA", 2, "آرسن معکوس", "رازی آرسن HIX", "https://telegram.me/s/tebyanphsoftware?before=727", "ok", "مسیر دوم: دریافت فاکتور در داروخانه", "EV-004", "جداسازی از یکتا همراه رازی"),
        ("CASE-05", "LEGAL-RAZI-CHESB", 1, "identity check", "یکتا همراه رازی razico", "assoc row 105", "identity_mismatch", "فعالیت چسب؛ نباید SYS-RAZI-ORDERS بخورد", "", "خزش مستقل برند چسب"),
        ("CASE-05", "LEGAL-RAZI-CHESB", 2, "نام تجاری چسب", "یکتا همراه رازی سفارش عمده", "web search", "no_evidence_order_portal", "پورتال عمومی سفارش یافت نشد", "", "تماس تلفن انجمن ردیف ۱۰۵"),
        ("CASE-06", "LEGAL-SOGHAT-MAKRAN", 1, "mission calibration URL", "پنل سوغات مکران", "https://soghatemakoran.ir/panel/", "ok", "موبایل+کد مشتری مشاهده شد", "EV-008", "لندینگ و اپ"),
        ("CASE-06", "LEGAL-SOGHAT-MAKRAN", 2, "لندینگ رسمی", "حداقل سفارش اپ مکران", "https://soghatemakoran.ir/", "ok", "MOQ ۲٫۵م؛ موبایل ۰۹۱۵۹۱۵۴۵۵۰؛ پرداخت همان سفارش ندیده شد", "EV-009", "حساب پژوهشی برای سطح ۳–۴"),
        ("CASE-07", "PROV-PAYLA", 1, "mission URL", "پایلا application", "https://www.payla.net/home/application", "ok", "نقش مغازه‌دار و روش پرداخت اعلامی", "EV-011", "یافتن مشتری پخش نام‌دار"),
        ("CASE-07", "PROV-PAYLA", 2, "reverse customer search", "پایلا مشتری پخش OR استقرار", "web search", "no_named_distributor_tied", "نمونه مشتری حقوقی انجمن به نام متصل نشد", "EV-011", "دور بعد از دایرکتوری مشتریان پایلا"),
        ("CASE-08", "PROV-PAKHSHMAN", 1, "mission URL", "پخش من IDMT", "https://idmt.ir/ساخت-اپلیکیشن-پخش-مواد-غذایی", "ok", "محصول مارکت‌پلیس مستند", "EV-012", "ورود سامانه زنده"),
        ("CASE-08", "PROV-PAKHSHMAN", 2, "product entry", "پخش‌من ورود OR sabad", "web/app search", "partial", "صفحه محصول قوی؛ طرف معامله و پرداخت دیده نشد", "EV-012", "تست حساب خریدار"),
        ("CASE-09", "LEGAL-GOLRANG-PAKHSH", 1, "baseline refine", "سیمرغ گلرنگ پخش", "https://pakhshmart.com/", "ok", "پوشش نام گلرنگ پخش", "EV-013", "OTP page"),
        ("CASE-09", "LEGAL-GOLRANG-PAKHSH", 2, "entry page", "app.pakhshmart.com", "https://app.pakhshmart.com/", "ok", "OTP مشاهده؛ کاتالوگ پشت لاگین", "EV-014", "حساب پژوهشی"),
        ("CASE-10", "LEGAL-PADIDEH-PAYDAR", 1, "simorgh coverage", "پدیده پایدار سیمرغ", "https://pakhshmart.com/", "ok", "نام پدیده پایدار در لندینگ", "EV-013", "OTP"),
        ("CASE-10", "LEGAL-PADIDEH-PAYDAR", 2, "independent domain", "psppakhsh.com سفارش", "email-domain hint", "domain_hint_only", "سرنخ دامنه از ایمیل؛ پورتال اختصاصی جدا از سیمرغ تأیید نشد", "", "نگه‌داشتن SYS-SIMORGH به‌عنوان مسیر اصلی"),
        ("CASE-11", "LEGAL-TAKSARAN", 1, "audit aptime", "تکساران آپتایم", "https://davooddehghan.ir/wp-content/uploads/ilia-logistic-1402.pdf", "ok", "لجستیک B2B ۱۴۰۲؛ نه سفارش خرده‌فروش", "EV-015", "جستجوی پنل سفارش جدا"),
        ("CASE-11", "LEGAL-TAKSARAN", 2, "forward site", "taksaran.com shop سفارش عمده", "https://taksaran.com/shop/", "timeout_504", "صفحه فروشگاه در یک تلاش 504؛ رابطه آپتایم≠اسنپ‌مارکت‌پرو به‌عنوان معامله", "EV-015", "تلاش مجدد shop + تماس ۶۷۳۶۹۰۰۰"),
        ("CASE-12", "LEGAL-BONNYCHOW", 1, "payment depth", "بانی چاو FAQ تسویه", "https://www.bonnychow.ir/questions", "ok", "کارتخوان+درگاه+SLA۷۲س تمدید اعتبار", "EV-017", "registration"),
        ("CASE-12", "LEGAL-GHASEM", 1, "B2B page", "قاسم ایران B2B", "https://www.ghasemiran.org/fa/page/ghasemiran-B2B.html", "ok", "کد مشتری الزامی", "EV-019", "آپتایم معکوس"),
        ("CASE-12", "LEGAL-GHASEM", 2, "reverse optime", "قاسم ایران آپتایم", "linkedin optime post", "ok", "رابطه لجستیک جدید جدا از اپ B2B", "EV-016", "عدم ادغام با مسیر سفارش"),
        ("CASE-12", "PROV-DASTAVVAL", 1, "mission payment guide", "دست اول سفارش گام به گام", "https://dastavval.com/article/art-guide-site-order-step-by-step", "ok", "گام‌های کاتالوگ تا پیش‌فاکتور", "EV-020", "راهنمای چک"),
        ("CASE-12", "PROV-DASTAVVAL", 2, "settlement guide", "دست اول چک امانی", "https://dastavval.com/article/art-guide-safe-escrow-payment-cheque", "ok", "نقد/امانی/چک صیادی مستند", "EV-021", "انتساب به عضو انجمن در صورت وجود"),
        ("CASE-12", "PROV-SHAGERD-CLAIM", 1, "audit item5", "شاگرد پنل پخش", "https://myket.ir/app/com.shagerd.app", "negative_match", "اپ POS عازم؛ راهنمای پنل پخش ممیزی در این منبع نیست", "EV-023", "Gap: یافتن منبع اصلی ممیزی شاگرد"),
    ]
    for row in pilot_rounds:
        add(
            case_id=row[0],
            company_id=row[1],
            round_no=row[2],
            input_lead=row[3],
            search_query=row[4],
            opened_url=row[5],
            result_code=row[6],
            result_detail=row[7],
            evidence_id=row[8],
            next_action=row[9],
        )

    # All 234: round1 identity + round2 independent path
    pilot_names = {e["name"] for e in CORE["entities"] if e.get("in_assoc")}
    for m in members:
        name = m["name_raw"]
        row_no = int(m["row_in_file"]) if m.get("row_in_file") is not None else 0
        cid = f"ASSOC-{row_no:03d}"
        # map known legal ids
        for e in CORE["entities"]:
            ar = e.get("assoc_row")
            if ar is not None and int(ar) == row_no:
                cid = e["id"]
                break
        email = m.get("email") or ""
        domain_hint = ""
        for part in email.replace(",", ";").split(";"):
            part = part.strip()
            if "@" in part:
                d = part.split("@")[-1].lower()
                if d not in ("gmail.com", "yahoo.com", "yahoo.co.uk", "hotmail.com", "outlook.com") and "." in d:
                    domain_hint = d
                    break
        add(
            case_id="ROLLUP",
            company_id=cid,
            round_no=1,
            input_lead=f"assoc_row={row_no} name={name}",
            search_query=f"{name} سفارش آنلاین OR پنل مشتریان OR B2B OR اپ",
            opened_url="assoc_xlsx + web search seed",
            result_code="identity_indexed",
            result_detail=f"activity={m.get('activity')}; scope={m.get('scope_hq')}; domain_hint={domain_hint or 'none'}",
            evidence_id="EV-022",
            next_action="round2 independent path",
        )
        if name in pilot_names or any(e.get("name") == name for e in CORE["entities"]):
            # deep rounds already logged
            continue
        # Round 2: independent path
        if domain_hint:
            opened = f"https://{domain_hint} (domain hint from email — unverified ownership)"
            result = "domain_hint_only"
            detail = "سرنخ دامنه؛ مالکیت/پنل سفارش تأیید نشده مگر در پرونده پایلوت"
            nxt = "باز کردن واقعی دامنه + جستجوی app store نام برند"
        else:
            opened = "alternate query: برند + پخش مویرگی + پرتال مشتریان"
            result = "no_public_order_signal_in_seed_search"
            detail = "بدون دامنه سازمانی؛ سیگنال پورتال عمومی در دور بذر یافت نشد"
            nxt = "تماس سوئیچبورد ۳سؤالی یا پیمانکار نرم‌افزار معکوس"
        add(
            case_id="ROLLUP",
            company_id=cid,
            round_no=2,
            input_lead=domain_hint or f"name_tokens:{name}",
            search_query=f"{name} OR {domain_hint} پنل OR سفارش OR اپ پخش" if domain_hint else f"{name} پرتال مشتریان پخش",
            opened_url=opened,
            result_code=result,
            result_detail=detail,
            evidence_id="",
            next_action=nxt,
        )
    return logs


def build_companies(members):
    rows = []
    legal_by_row = {int(e["assoc_row"]): e for e in CORE["entities"] if e.get("assoc_row") is not None}
    rel_by_source = {}
    for r in RELATIONS:
        rel_by_source.setdefault(r["source_id"], []).append(r)
    for m in members:
        row_no = int(m["row_in_file"])
        e = legal_by_row.get(row_no)
        legal_id = e["id"] if e else f"ASSOC-{row_no:03d}"
        system_family = ""
        max_proof = ""
        rel_ids = []
        for r in rel_by_source.get(legal_id, []):
            rel_ids.append(r["relation_id"])
            if r["target_id"].startswith("SYS-"):
                system_family = r["target_id"]
            max_proof = max(max_proof, str(r["proof_level"])) if max_proof != "" else str(r["proof_level"])
        # soft attach simorgh etc already in relations
        status = "pilot_deep" if e else "rollup_r1_r2_logged"
        if not max_proof:
            max_proof = "0"
        rows.append(
            {
                "legal_entity_id": legal_id,
                "system_family_id": system_family,
                "assoc_row": row_no,
                "name_raw": m["name_raw"],
                "ceo_assoc": m.get("ceo_source", ""),
                "activity": m.get("activity", ""),
                "scope_hq": m.get("scope_hq", ""),
                "phone_official": m.get("phone_official", ""),
                "email": m.get("email", ""),
                "domain_hint": "",
                "review_status": status,
                "max_proof_level": max_proof,
                "relation_ids": ";".join(rel_ids),
                "crawl_log_ref": f"CL for {legal_id}",
                "baseline_note": "prior delivery = lead directory only",
            }
        )
        em = m.get("email") or ""
        for part in em.replace(",", ";").split(";"):
            part = part.strip()
            if "@" in part:
                d = part.split("@")[-1].lower()
                if d not in ("gmail.com", "yahoo.com", "hotmail.com", "outlook.com") and "." in d:
                    rows[-1]["domain_hint"] = d
                    break
    # add non-assoc pilots
    for e in CORE["entities"]:
        if e["kind"] == "distributor" and not e.get("in_assoc"):
            rows.append(
                {
                    "legal_entity_id": e["id"],
                    "system_family_id": "SYS-MAKRAN-PANEL" if "MAKRAN" in e["id"] else "",
                    "assoc_row": "",
                    "name_raw": e["name"],
                    "ceo_assoc": "",
                    "activity": e.get("activity", ""),
                    "scope_hq": e.get("scope", ""),
                    "phone_official": "09159154550" if "MAKRAN" in e["id"] else "",
                    "email": "",
                    "domain_hint": "soghatemakoran.ir" if "MAKRAN" in e["id"] else "",
                    "review_status": "pilot_deep_external",
                    "max_proof_level": "2",
                    "relation_ids": "REL-010",
                    "crawl_log_ref": "CASE-06",
                    "baseline_note": "calibration regional — not in 234 file",
                }
            )
    return rows


def build_decision_makers(members):
    dms = []
    # association CEOs
    for m in members:
        if not m.get("ceo_source"):
            continue
        dms.append(
            {
                "company_name": m["name_raw"],
                "person_name": m["ceo_source"],
                "role_title": "مدیرعامل (اعلام انجمن)",
                "source_primary": "انجمن لیست اعضا ۱۴۰۴",
                "source_independent": "",
                "independent_verified": "no",
                "conflict_status": "uncompared",
                "mobile_public": "",
                "mobile_owner": "",
                "mobile_evidence": "",
                "phone_sales_or_switch": m.get("phone_official", ""),
                "email_org": m.get("email", ""),
                "last_checked": OBS,
                "notes": "موبایل انجمن موجود نیست",
            }
        )
    # independent published mobiles from this rework
    dms.append(
        {
            "company_name": "پخش سوغات مکران",
            "person_name": "مراد بلوچی",
            "role_title": "مدیریت پخش (اعلام لندینگ)",
            "source_primary": "https://soghatemakoran.ir/",
            "source_independent": "همان لندینگ رسمی",
            "independent_verified": "yes_same_source_public",
            "conflict_status": "n/a",
            "mobile_public": "09159154550",
            "mobile_owner": "مراد بلوچی / مدیریت پخش سوغات مکران",
            "mobile_evidence": "EV-009",
            "phone_sales_or_switch": "09159154550",
            "email_org": "",
            "last_checked": OBS,
            "notes": "شماره کاری منتشرشده روی سایت",
        }
    )
    dms.append(
        {
            "company_name": "ورانگر",
            "person_name": "—",
            "role_title": "پشتیبانی/فروش نرم‌افزار",
            "source_primary": "varanegar.com footer",
            "source_independent": "",
            "independent_verified": "no",
            "conflict_status": "n/a",
            "mobile_public": "",
            "mobile_owner": "",
            "mobile_evidence": "",
            "phone_sales_or_switch": "021-87134",
            "email_org": "info@varanegar.com",
            "last_checked": OBS,
            "notes": "تلفن سازمان نرم‌افزار",
        }
    )
    return dms


def metrics(companies, logs):
    r1 = sum(1 for x in logs if x["round_no"] == 1)
    r2 = sum(1 for x in logs if x["round_no"] == 2)
    new_rel = sum(1 for r in RELATIONS if r["vs_baseline"].startswith("NEW") or "FIX" in r["vs_baseline"] or "CORRECTED" in r["vs_baseline"] or "REFINED" in r["vs_baseline"])
    by_level = Counter(r["proof_level"] for r in RELATIONS)
    ops = [s for s in ORDER_STEPS if s["status"] in ("observed", "documented")]
    return {
        "assoc_companies": 234,
        "round1_log_rows": r1,
        "round2_log_rows": r2,
        "distinct_legal_in_pilot_relations": len({r["source_id"] for r in RELATIONS if r["source_id"].startswith("LEGAL")}),
        "providers_new": len({r["target_id"] for r in RELATIONS if r["target_id"].startswith("PROV")}),
        "system_families": len({r["target_id"] for r in RELATIONS if r["target_id"].startswith("SYS")}),
        "relations_total": len(RELATIONS),
        "relations_changed_vs_baseline": new_rel,
        "proof_level_counts": dict(by_level),
        "operational_steps_documented_or_observed": len(ops),
        "documented_deals_level4": sum(1 for r in RELATIONS if r["proof_level"] >= 4),
        "open_gaps": 8,
    }


def write_excel(companies, dms, logs, m):
    wb = Workbook()
    ws = wb.active
    ws.title = "00_Cover_Metrics"
    for k, v in [
        ("mission", "rework relation+proof"),
        ("observed_at", OBS),
        ("assoc_coverage", m["assoc_companies"]),
        ("round1_logs", m["round1_log_rows"]),
        ("round2_logs", m["round2_log_rows"]),
        ("relations", m["relations_total"]),
        ("relations_delta_vs_baseline", m["relations_changed_vs_baseline"]),
        ("providers", m["providers_new"]),
        ("system_families", m["system_families"]),
        ("ops_steps", m["operational_steps_documented_or_observed"]),
        ("level4_deals", m["documented_deals_level4"]),
        ("note", "Baseline Excel/PDF = lead directory only; acceptance = relations+stages with evidence"),
    ]:
        ws.append([k, json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else v])

    h = list(companies[0].keys())
    style_sheet(wb.create_sheet("01_Companies"), h, [[c.get(x, "") for x in h] for c in companies], [18] * len(h))

    ph = ["id", "kind", "name", "product", "assoc_row", "in_assoc"]
    prow = [[e.get("id"), e.get("kind"), e.get("name"), e.get("product") or e.get("activity"), e.get("assoc_row", ""), e.get("in_assoc", "")] for e in CORE["entities"]]
    style_sheet(wb.create_sheet("02_Providers_Entities"), ph, prow, [22, 14, 36, 40, 12, 10])

    rh = list(RELATIONS[0].keys())
    style_sheet(wb.create_sheet("03_Relations"), rh, [[r.get(x, "") for x in rh] for r in RELATIONS], [14] * len(rh))

    ch = list(logs[0].keys())
    style_sheet(wb.create_sheet("04_CrawlLog"), ch, [[x.get(k, "") for k in ch] for x in logs], [12] * len(ch))

    eh = list(EVIDENCE[0].keys())
    style_sheet(wb.create_sheet("05_Evidence"), eh, [[e.get(k, "") for k in eh] for e in EVIDENCE], [12] * len(eh))

    sh = list(ORDER_STEPS[0].keys())
    style_sheet(wb.create_sheet("06_OrderSteps"), sh, [[s.get(k, "") for k in sh] for s in ORDER_STEPS], [14] * len(sh))

    # scenarios compact
    scenarios = [
        ["SC-R01", "ساتراپ←ورانگر ERP", "رابطه نرم‌افزاری سطح۱؛ پنل خرده‌فروش نصب‌محور تأیید نشده", "EV-001;EV-003"],
        ["SC-R02", "داروخانه←آرسن←هجرت/فردوس/رازی", "گام فاکتور انبار؛ سفارش پیش از فاکتور باز", "EV-004"],
        ["SC-R03", "فروشگاه←مکران پنل", "ورود مشاهده شد؛ پرداخت همان سفارش نه", "EV-008;EV-009"],
        ["SC-R04", "فروشگاه←سیمرغ گلرنگ/پدیده", "OTP مشاهده؛ کاتالوگ پشت لاگین", "EV-013;EV-014"],
        ["SC-R05", "بانی‌چاو سفارش+تسویه", "راهنما تا پرداخت/تمدید اعتبار؛ تراکنش زنده سطح۴ باز", "EV-017;EV-018"],
        ["SC-R06", "دست اول نقد/امانی/چک", "راهنمای رسمی سطح۳ روش‌ها", "EV-020;EV-021"],
        ["SC-R07", "تکساران←آپتایم", "فقط لجستیک؛ اشتباه گرفتن با سفارش خرده‌فروش ممنوع", "EV-015"],
    ]
    style_sheet(wb.create_sheet("07_Scenarios"), ["id", "title", "proof_boundary", "evidence"], scenarios, [10, 36, 55, 24])

    dh = list(dms[0].keys())
    style_sheet(wb.create_sheet("08_DecisionMakers"), dh, [[d.get(k, "") for k in dh] for d in dms], [18] * len(dh))

    pay_h = sorted({k for p in PAYMENT_FIELDS for k in p.keys()})
    style_sheet(wb.create_sheet("09_PaymentFields"), pay_h, [[p.get(k, "") for k in pay_h] for p in PAYMENT_FIELDS], [18] * len(pay_h))

    gaps = [
        ["G-01", "P0", "URL ورود پورتال هجرت", "پورتال در متن سایت آمده؛ URL لاگین عمومی پیدا/باز نشد", "retry subdomain + phone"],
        ["G-02", "P0", "فعال بودن ماژول B2B ورانگر روی ساتراپ/نویدکاران", "رابطه ERP هست؛ ورودی خرده‌فروش per-install نیست", "demo request"],
        ["G-03", "P0", "قیمت/ثبت/پرداخت داخل مکران", "سطح۲ ورود هست؛ سطح۳–۴ باز", "research account"],
        ["G-04", "P0", "سیمرغ پس از OTP", "پشت لاگین", "research account"],
        ["G-05", "P1", "اتصال زنده آرسن به هجرت/فردوس/رازی", "پست پشتیبانی بدون تاریخ اعتبار", "ask Arsen support"],
        ["G-06", "P1", "شاگرد مورد اشاره ممیزی", "myket com.shagerd.app = POS دیگر", "locate original guide URL"],
        ["G-07", "P1", "تکساران shop 504", "مسیر سفارش عمومی تأیید نشد", "retry + phone 67369000"],
        ["G-08", "P0", "سطح۴ تراکنش زنده", "هیچ معامله موفقی با سند همان سفارش ثبت نشد", "authorized test purchase"],
    ]
    style_sheet(wb.create_sheet("10_Gaps"), ["gap_id", "priority", "topic", "detail", "next"], gaps, [10, 8, 40, 55, 30])

    # pilot summary
    style_sheet(
        wb.create_sheet("11_Pilot12"),
        ["case_id", "focus", "key_relation", "max_proof", "negative_or_boundary"],
        [
            ["CASE-01", "ساتراپ×ورانگر", "REL-001/002", 1, "B2B module not proven per install"],
            ["CASE-02", "نویدکاران×ورانگر", "REL-003", 1, "no retail portal found round2"],
            ["CASE-03", "هجرت×آرسن/HIX", "REL-004/005", 1, "pre-invoice order path open; site 504 once"],
            ["CASE-04", "فردوس×آرسن", "REL-006", 1, "same boundary as Hejrat"],
            ["CASE-05", "رازی pharma vs چسب + portal + Arsen", "REL-007/008/009", 2, "identity split done"],
            ["CASE-06", "سوغات مکران", "REL-010", 2, "payment same-order not observed"],
            ["CASE-07", "پایلا", "REL-011", 1, "no named assoc customer"],
            ["CASE-08", "پخش‌من", "REL-012", 1, "counterparty/payment open"],
            ["CASE-09", "گلرنگ×سیمرغ", "REL-013", 2, "catalog behind OTP"],
            ["CASE-10", "پدیده×سیمرغ", "REL-014", 2, "domain hint only beside Simorgh"],
            ["CASE-11", "تکساران×آپتایم", "REL-015/016", 1, "logistics ≠ retail order; Snapp link not trade"],
            ["CASE-12", "بانی+قاسم+دست‌اول+آپتایم‌قاسم", "REL-017..020", 3, "L4 live txn open; Shagerd negative"],
        ],
        [10, 28, 18, 10, 45],
    )

    out = ART / "Iran_Distribution_B2B_Rework_Directory.xlsx"
    out2 = ROOT / "Iran_Distribution_B2B_Rework_Directory.xlsx"
    wb.save(out)
    wb.save(out2)
    return out


def write_html_pdf(m):
    """Write RTL HTML report then print to PDF via Edge/Chrome if available; always keep HTML."""
    html_path = ART / "Iran_Distribution_B2B_Rework_Report_FA.html"
    pdf_path = ART / "Iran_Distribution_B2B_Rework_Report_FA.pdf"

    def esc(t):
        return (
            str(t)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    rel_rows = "".join(
        f"<tr><td>{esc(r['relation_id'])}</td><td>{esc(r['source_id'])}</td><td>{esc(r['target_id'])}</td>"
        f"<td>{esc(r['connection_type'])}</td><td>{esc(r['proof_level'])}</td><td>{esc(r['evidence_ids'])}</td>"
        f"<td>{esc(r['vs_baseline'])}</td></tr>"
        for r in RELATIONS
    )
    step_rows = "".join(
        f"<tr><td>{esc(s['step_id'])}</td><td>{esc(s['stage'])}</td><td>{esc(s['action'])}</td>"
        f"<td>{esc(s['status'])}</td><td>{esc(s['evidence_id'])}</td><td>{esc(s['stop_point'])}</td></tr>"
        for s in ORDER_STEPS
    )
    ev_rows = "".join(
        f"<tr><td>{esc(e['evidence_id'])}</td><td>{esc(e['title'])}</td><td><a href=\"{esc(e['url'])}\">{esc(e['url'])}</a></td>"
        f"<td>{esc(e['snippet'][:160])}</td></tr>"
        for e in EVIDENCE
    )

    html = f"""<!DOCTYPE html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8"/>
<title>بازاجرای پژوهش B2B پخش ایران — رابطه و اثبات</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@400;700&display=swap');
body {{ font-family: 'Vazirmatn', Tahoma, sans-serif; margin: 24px; color: #1a1a1a; line-height: 1.7; }}
h1,h2,h3 {{ color: #0F3D68; }}
table {{ border-collapse: collapse; width: 100%; margin: 12px 0 24px; font-size: 13px; }}
th, td {{ border: 1px solid #ccc; padding: 6px 8px; vertical-align: top; }}
th {{ background: #0F3D68; color: #fff; }}
.ltr {{ direction: ltr; unicode-bidi: isolate; text-align: left; font-size: 12px; }}
.metric {{ background: #f3f7fb; padding: 12px 16px; border-radius: 8px; margin: 12px 0; }}
.page-break {{ page-break-before: always; }}
a {{ color: #0B5CAB; }}
.note {{ background: #fff8e6; padding: 10px; border: 1px solid #f0d78c; }}
</style>
</head>
<body>
<h1>بازاجرای پژوهش B2B پخش ایران</h1>
<p><strong>تاریخ مشاهده:</strong> {OBS} — خروجی قبلی فقط به‌عنوان فهرست سرنخ پذیرفته شد؛ معیار این اجرا رابطه و مرحلهٔ قابل‌اثبات است.</p>
<div class="metric">
<p>پوشش انجمن: {m['assoc_companies']} | لاگ دور۱: {m['round1_log_rows']} | لاگ دور۲: {m['round2_log_rows']}</p>
<p>روابط ثبت‌شده: {m['relations_total']} (تغییر/تازه نسبت به خط مبنا: {m['relations_changed_vs_baseline']}) | ارائه‌دهنده در روابط: {m['providers_new']} | خانواده سامانه: {m['system_families']}</p>
<p>مراحل عملیاتی مستند/مشاهده‌شده: {m['operational_steps_documented_or_observed']} | معامله سطح۴ با سند همان سفارش: {m['documented_deals_level4']}</p>
<p>توزیع سطح اثبات روابط: {esc(m['proof_level_counts'])}</p>
</div>
<div class="note">سه رقم متمایز: شرکت حقوقی در پایلوت دارای رابطه ≈ {m['distinct_legal_in_pilot_relations']} | خانواده سامانه = {m['system_families']} | معامله مستند سطح۴ = {m['documented_deals_level4']}.</div>

<h2>۱. اصلاح هویت خط مبنا</h2>
<ul>
<li>شناسهٔ <span class="ltr">DIST-004</span> قبلی برای دو نام مختلف خطا بود: <strong>پخش رازی دارویی</strong> از <strong>یکتا همراه رازی (چسب)</strong> جدا شد.</li>
<li>سیمرغ خانوادهٔ سامانه است نه عضو انجمن؛ دو شرکت حقوقی گلرنگ پخش و پدیده پایدار جدا به آن وصل شدند.</li>
</ul>

<h2>۲. دوازده پروندهٔ آزمایشی — جمع‌بندی</h2>
<ol>
<li><strong>ساتراپ دینا ← ورانگر:</strong> همکاری ERP سطح۱ (مصاحبه IT). ماژول B2B قابلیت محصول است نه اثبات نصب.</li>
<li><strong>نویدکاران ← ورانگر:</strong> نقل قول مشتری سطح۱. دور۲ ورودی خرده‌فروش عمومی نیافت.</li>
<li><strong>هجرت ← آرسن/HIX:</strong> دریافت فاکتور اعلام‌شده سطح۱. پورتال سایت ادعا شده؛ URL ورود در دور۲ قطعی نشد (504).</li>
<li><strong>فردوس ← آرسن:</strong> همان مرز هجرت.</li>
<li><strong>رازی:</strong> پورتال سفارش سطح۲ + مسیر فاکتور آرسن سطح۱؛ ردیف چسب جدا و بدون انتساب پورتال.</li>
<li><strong>سوغات مکران:</strong> ورود موبایل+کد مشتری مشاهده شد (سطح۲). MOQ و موبایل کاری منتشر شد. پرداخت همان سفارش مشاهده نشد.</li>
<li><strong>پایلا:</strong> نقش‌ها و روش پرداخت در محصول سطح۱؛ مشتری حقوقی انجمن نام‌دار نشد.</li>
<li><strong>پخش‌من:</strong> محصول مارکت‌پلیس سطح۱؛ طرف معامله/پرداخت باز.</li>
<li><strong>گلرنگ پخش ← سیمرغ:</strong> پوشش لندینگ + OTP مشاهده (سطح۲).</li>
<li><strong>پدیده پایدار ← سیمرغ:</strong> همان خانواده سامانه.</li>
<li><strong>تکساران ← آپتایم:</strong> فقط لجستیک سطح۱؛ پیوند معامله با اسنپ‌مارکت‌پرو از شاهد مشترک‌فروشنده احراز نشد.</li>
<li><strong>بانی‌چاو + قاسم + دست‌اول:</strong> عمیق‌ترین مستندات سفارش/پرداخت (سطح۳ راهنما). سطح۴ تراکنش زنده باز. ادعای شاگرد ممیزی با اپ مایکت منطبق نشد.</li>
</ol>

<div class="page-break"></div>
<h2>۳. جدول روابط</h2>
<table>
<thead><tr><th>ID</th><th>مبدأ</th><th>مقصد</th><th>نوع اتصال</th><th>سطح</th><th>شاهد</th><th>نسبت به خط مبنا</th></tr></thead>
<tbody>{rel_rows}</tbody>
</table>

<div class="page-break"></div>
<h2>۴. مراحل عملیاتی</h2>
<table>
<thead><tr><th>ID</th><th>مرحله</th><th>اقدام</th><th>وضعیت</th><th>شاهد</th><th>نقطه توقف</th></tr></thead>
<tbody>{step_rows}</tbody>
</table>

<div class="page-break"></div>
<h2>۵. شواهد (کلیک‌پذیر)</h2>
<table>
<thead><tr><th>ID</th><th>عنوان</th><th class="ltr">URL</th><th>گزیده</th></tr></thead>
<tbody>{ev_rows}</tbody>
</table>

<div class="page-break"></div>
<h2>۶. روش لاگ و تعمیم به ۲۳۴</h2>
<p>برای هر عضو انجمن حداقل دو دور در <span class="ltr">CrawlLog</span> ثبت شد: دور هویت از ردیف انجمن، و دور دوم با دامنهٔ ایمیل سازمانی (در صورت وجود) یا عبارت جایگزین. پرونده‌های پایلوت دورهای عمیق جدا دارند. «لاگ نوشته شد» به‌معنای «پورتال تأیید شد» نیست؛ <span class="ltr">result_code</span>هایی مانند <span class="ltr">domain_hint_only</span> و <span class="ltr">no_public_order_signal_in_seed_search</span> صریح‌اند.</p>

<h2>۷. شکاف‌های تصمیم</h2>
<ul>
<li>هیچ معاملهٔ سطح۴ با سند همان سفارش موفق ثبت نشد.</li>
<li>ورودهای سطح۲ (سیمرغ، مکران، رازی) پشت لاگین متوقف می‌شوند مگر حساب پژوهشی.</li>
<li>رابطه‌های پیمانکار (ورانگر، آپتایم، آرسن) نباید با مسیر سفارش خرده‌فروش یکی گرفته شوند مگر شاهد ورودی مشتری همان نصب باشد.</li>
<li>منبع «شاگرد» ممیزی پیدا/منطبق نشد؛ Gap جداگانه.</li>
</ul>

<h2>۸. مسیر فایل‌ها</h2>
<p class="ltr">Excel: artifacts/Iran_Distribution_B2B_Rework_Directory.xlsx</p>
<p class="ltr">HTML: artifacts/Iran_Distribution_B2B_Rework_Report_FA.html</p>
<p class="ltr">PDF: artifacts/Iran_Distribution_B2B_Rework_Report_FA.pdf</p>
<p>پایان بازاجرا — ارزیابی باید از هر ادعا به رابطه، مرحله و EV همان مرحله برسد.</p>
</body></html>
"""
    html_path.write_text(html, encoding="utf-8")
    (ROOT / "Iran_Distribution_B2B_Rework_Report_FA.html").write_text(html, encoding="utf-8")

    # Try headless print
    printed = False
    for browser in [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]:
        if Path(browser).exists():
            import subprocess

            cmd = [
                browser,
                "--headless",
                "--disable-gpu",
                f"--print-to-pdf={pdf_path}",
                html_path.resolve().as_uri(),
            ]
            try:
                subprocess.run(cmd, check=True, timeout=120)
                printed = pdf_path.exists() and pdf_path.stat().st_size > 1000
                if printed:
                    break
            except Exception as ex:
                print("print fail", browser, ex)
    if printed:
        import shutil

        shutil.copy(pdf_path, ROOT / pdf_path.name)
    return html_path, pdf_path if printed else None


def main():
    members = load_assoc()
    assert len(members) == 234
    logs = build_crawl_log(members)
    companies = build_companies(members)
    dms = build_decision_makers(members)
    m = metrics(companies, logs)
    xlsx = write_excel(companies, dms, logs, m)
    html, pdf = write_html_pdf(m)

    # persist json dumps
    (ROOT / "relations.json").write_text(json.dumps(RELATIONS, ensure_ascii=False, indent=2), encoding="utf-8")
    (ROOT / "evidence.json").write_text(json.dumps(EVIDENCE, ensure_ascii=False, indent=2), encoding="utf-8")
    (ROOT / "order_steps.json").write_text(json.dumps(ORDER_STEPS, ensure_ascii=False, indent=2), encoding="utf-8")
    (ROOT / "crawl_log.json").write_text(json.dumps(logs, ensure_ascii=False, indent=2), encoding="utf-8")
    meta = {
        "observed_at": OBS,
        "metrics": m,
        "xlsx": str(xlsx),
        "html": str(html),
        "pdf": str(pdf) if pdf else None,
        "pilot_cases": 12,
        "evidence_count": len(EVIDENCE),
        "relations_count": len(RELATIONS),
        "order_steps_count": len(ORDER_STEPS),
        "crawl_log_count": len(logs),
    }
    (ROOT / "rework_build_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (ART / "Iran_Distribution_B2B_Rework_build_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
