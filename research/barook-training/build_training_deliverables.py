# -*- coding: utf-8 -*-
"""Build Barook training-flow extraction deliverables."""
from __future__ import annotations
import csv, json, shutil
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parent
ART = Path(r"D:\HomeDesign\artifacts\barook-training")
ART.mkdir(parents=True, exist_ok=True)
TRANS = ROOT / "transcripts"
SHOT = ROOT / "screenshots"
TRANS.mkdir(exist_ok=True)
SHOT.mkdir(exist_ok=True)

OBS = "1405/09/04"
ACCESS_BARRIER = (
    "بایت ویدئو از این ماشین پژوهشی قابل دریافت نبود: "
    "اتصال مستقیم به t.me و cdn*.telesco.pe با timeout قطع می‌شود "
    "(Playwright/Chrome و SOCKS5 هم ناموفق). "
    "کپشن رسمی و مدت از ویجت تلگرام/WebFetch خوانده شد. "
    "فریم داخل‌ویدئو تولید یا جعل نشده است."
)

SOURCES = [
    dict(id="SRC-TG-055", title="فهرست آموزش‌های باروک (ایندکس)", role="all", product="Barook trainings index",
         url="https://t.me/barook_ir/55", publisher="barook_ir", published="observed via embed", duration="",
         access="caption_ok_video_blocked", transcript="", flows="index"),
    dict(id="SRC-TG-110", title="پرداخت آنلاین توسط خریدار", role="supermarket", product="online credit pay",
         url="https://t.me/barook_ir/110", publisher="barook_ir", published="embed observed", duration="0:52",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-110.md", flows="F01"),
    dict(id="SRC-TG-108", title="ثبت سفارش فروشگاه توسط شرکت پخش (آنلاین اوردرینگ)", role="distributor", product="panel order register",
         url="https://t.me/barook_ir/108", publisher="barook_ir", published="embed observed", duration="0:35",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-108.md", flows="F01,F04"),
    dict(id="SRC-TG-122", title="پوز — پرداخت پیش‌فاکتورهای پرداخت‌نشده", role="distributor+supermarket", product="POS unpaid",
         url="https://t.me/barook_ir/122", publisher="barook_ir", published="embed observed", duration="0:34",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-122.md", flows="F02"),
    dict(id="SRC-TG-121", title="پوز — فروش اعتباری", role="distributor+supermarket", product="POS credit sale",
         url="https://t.me/barook_ir/121", publisher="barook_ir", published="embed observed", duration="0:42",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-121.md", flows="F03"),
    dict(id="SRC-TG-141", title="بارگذاری فاکتورها در پنل تأمین‌کننده", role="supplier", product="invoice upload",
         url="https://t.me/barook_ir/141", publisher="barook_ir", published="embed observed", duration="1:41",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-141.md", flows="F05"),
    dict(id="SRC-TG-150", title="شارژ کیف پول باروک", role="supermarket", product="wallet top-up",
         url="https://t.me/barook_ir/150", publisher="barook_ir", published="embed observed", duration="0:58",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-150.md", flows="F07"),
    dict(id="SRC-TG-013", title="فعالسازی باروک‌کارت", role="supermarket", product="card activation",
         url="https://t.me/barook_ir/13", publisher="barook_ir", published="2024-07-ish", duration="1:08",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-013.md", flows="F08"),
    dict(id="SRC-TG-058", title="دستگاه پوز P10 و نصب باروک", role="distributor", product="POS P10 install",
         url="https://t.me/barook_ir/58", publisher="barook_ir", published="embed observed", duration="3:25",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-058.md", flows="F09"),
    dict(id="SRC-TG-060", title="به‌روزرسانی اپلیکیشن باروک روی پوز", role="distributor", product="POS app update",
         url="https://t.me/barook_ir/60", publisher="barook_ir", published="embed observed", duration="1:07",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-060.md", flows="F10"),
    dict(id="SRC-TG-109", title="راهنمای استفاده از کسب‌نو حقوقی", role="legal_company", product="Kasbno legal",
         url="https://t.me/barook_ir/109", publisher="barook_ir", published="embed observed", duration="0:58",
         access="caption_ok_video_blocked", transcript="transcripts/SRC-TG-109.md", flows="F11"),
    dict(id="SRC-TG-017", title="فیلم آموزشی نحوه کار با دستگاه pos باروک (نسخهٔ قدیمی‌تر)", role="distributor+supermarket", product="POS legacy",
         url="https://t.me/barook_ir/17", publisher="barook_ir", published="2024-07-09", duration="0:52",
         access="caption_thin_video_blocked", transcript="transcripts/SRC-TG-017.md", flows="F03-legacy"),
    dict(id="SRC-TG-SEARCH", title="جست‌وجوی #آموزش در کانال", role="all", product="catalog",
         url="https://telegram.me/s/barook_ir?q=%23آموزش", publisher="barook_ir", published=OBS,
         duration="", access="list_ok", transcript="", flows="catalog"),
    dict(id="SRC-OFF-FMCG", title="صفحه رسمی اعتبار کسب‌وکاری سوپرمارکت", role="supermarket", product="Kasbno FMCG",
         url="https://barook.ir/fmcg", publisher="barook.ir", published=OBS, duration="",
         access="full_html_ok", transcript="", flows="F06"),
    dict(id="SRC-PARTNER-AZIMI", title="باروک کارت چیست؟ رضوان پخش عظیمی", role="supermarket", product="partner guide",
         url="https://rezvanpakhsh.com/what-is-barook-cart/", publisher="rezvanpakhsh", published="partner page",
         duration="", access="full_html_ok", transcript="", flows="F06,F03"),
    dict(id="SRC-TG-BONAK", title="بنک‌مارکت: دو مسیر پرداخت آنلاین یا کارت هنگام تحویل", role="supermarket", product="partner path",
         url="https://t.me/s/barook_ir/110", publisher="barook_ir context around 110+", published="channel feed",
         duration="", access="caption_ok", transcript="", flows="F01,F02"),
    dict(id="SRC-PRIOR-XLSX", title="Barook_Kasbno_Deep_Research_Directory.xlsx — Evidence/OrderSteps/Gaps", role="research", product="prior",
         url="artifacts/Barook_Kasbno_Deep_Research_Directory.xlsx", publisher="prior agent pack", published="1405/09/03",
         duration="", access="local", transcript="", flows="G-BK-06 follow-up"),
]

# Caption texts for transcripts
CAPTIONS = {
    "SRC-TG-110": """پرداخت آنلاین توسط خریدار
دارندگان باروک‌کارت بعد از ثبت سفارش کالاهای درخواستی‌شان از شرکت‌های پخش، می‌توانند به صورت آنلاین اقدام به پرداخت خرید اعتباری خود کنند.
فقط کافیست بعد از ثبت سفارش توسط شرکت پخش، مغازه‌دار مراحلی را که در ویدیو نشان داده شده است طی کند.
با این سرویس که به تازگی در پنل باروک توسعه یافته است، بدون نیاز به حضور در محل امکان خرید و فروش اعتباری مهیا شده است.""",
    "SRC-TG-108": """شرکت‌های پخش طرف قرارداد باروک می‌توانند به روشی که در این ویدیو شرح داده شده است، سفارش‌های فروشگاه‌داران را در پنل خود ثبت کنند تا بعد از تایید خرید اعتباری توسط مغازه‌دار، خرید و فروش اعتباری تکمیل شود.
با این سرویس که به تازگی در پنل باروک توسعه یافته است، بدون نیاز به حضور در محل خریدار یا فروشنده، امکان خرید و فروش اعتباری مهیا شده است.
#فروش_آنلاین #آنلاین_اوردرینگ""",
    "SRC-TG-122": """🔺نحوه کار با دستگاه POSباروک (پرداخت پیش‌فاکتورهای صادر شده)
در این ویدیو نحوه تسویه اعتبار بر بستر دستگاه POS توسط موزع آموزش داده شده.
وارد اپلیکیشن باروک روی دستگاه POS می‌شی.
در بخش سفارشات، روی گزینه «پرداخت نشده‌ها» کلیک می‌کنی و در بخش جزییات، روی گزینه پرداخت کلیک می‌کنی.
برای تکمیل مرحله‌ی پرداخت، باید از دارنده باروک‌کارت بخوای که کارتش رو روی دستگاه بکشه. و تمام!""",
    "SRC-TG-121": """🔺نحوه کار با دستگاه POS باروک(فروش اعتباری)
اپلیکیشن باروک در POS را باز کنید و روی فروش اعتباری کلیک کنید.
در قسمت فروش اعتباری مبلغ جمع فاکتور و مدت زمان سررسید رو از نماینده پخش بپرسید و وارد کنید و در ادامه روی دکمه فروش اعتباری کلیک کنید.
در نهایت باروک کارت را بکشید و رمز خودتون رو وارد کنید.""",
    "SRC-TG-141": """✅ آموزش بارگذاری فاکتورها در پنل تامین کننده
در این ویدیو، قدم‌به‌قدم یاد می‌گیرید چطور فاکتورهاتون رو در پنل تامین‌کننده باروک آپلود کنید. فقط چند کلیک و تمام!
#آموزش""",
    "SRC-TG-150": """✅چند کلیک تا شارژ کیف پول باروک!
۱. وارد صفحه‌ی اصلی اپلیکشن باروک بشید.
۲. روی دکمه کیف پول در پایین اپ کلیک کنید. بعدش بزنید روی شارژ کیف پول.
۳. اینجا دوتا کار می‌تونی بکنی. یا مبلغ پیش‌فرض رو انتخاب کنی یا خودت هر عددی که می خوای رو وارد کنی و بری به درگاه پرداخت.
۴. تمام! کیف پولت حالا شارژه و می‌تونی خریدهاتون انجام بدی.""",
    "SRC-TG-013": """💳 آموزش فعالسازی باروک‌ کارت
همه کسانی که کارت خودشون را دریافت کردند، می‌تونن با تماشای این ویدویوی آموزشی، باروک کارتشون را فعال کنند.
#باروک #کسبنو #آموزش""",
    "SRC-TG-058": """با دستگاه پوز P10 آشنا بشید
دستگاه‌های پوزی که در اختیار شرکای تجاری باروک قرار می‌گیرد، P10 از «شرکت توسن تکنو»ست.
این دستگاه‌ها تمام‌صفحه‌ست و با سیستم عامل اندروید کار می‌کند.
در این ویدیو نحوه نصب باروک روی این دستگاه پوز آموزش داده شده‌است. #آموزش""",
    "SRC-TG-060": """✅ به روزرسانی اپلیکیشن باروک
در این ویدیو نحوه‌ی به‌روز رسانی اپلیکیشن باروک روی دستگاه POS آموزش داده شده است. #آموزش""",
    "SRC-TG-109": """راهنمای استفاده از کسب‌نو حقوقی
شرایط دریافت باروک برای شرکت‌های حقوقی متقاضی باروک‌کارت که نحوه مدیریت و تغییر حق امضا متناسب با روزنامه رسمی همچنین مدیریت فاکتورهای صادر شده در بازارگاه اعتباری باروک در این ویدو آموزش داده شده است.""",
    "SRC-TG-017": """فیلم آموزشی نحوه کار با دستگاه pos باروک
(کپشن کوتاه؛ جزئیات اجرایی در کپشن نیست — نسخهٔ قدیمی‌تر نسبت به ۱۲۱/۱۲۲)""",
}

REG_CAPTION_FROM_SEARCH = """وارد سایت باروک شوید و ثبت نام کنید.
۲. مدارک لازم: اطلاعات هویتی؛ اطلاعات فروشگاه؛ اسناد و مدارک
۳. نحوه انجام اعتبارسنجی: روی «مراحل دریافت اعتبار را تکمیل کنید» کلیک کن.
بعد از گذشت چند ساعت نتیجه اعتبارسنجی توی همون قسمت مشخص میشه. روی «ادامه مراحل بانکی» کلیک کنید.
۴. شعبه بانکی را انتخاب کنید: وقتی «در انتظار تعیین نوبت هستید»، ضربدر بالای صفحه را بزنید.
وارد صفحه اصلی که بشوید نام شعبه انتخابی و تاریخ مراجعه به بانک نمایش داده می‌شود.
۵. بعد از مراجعه به بانک با مدارک، منتظر اعتبارسنجی بانک باشید؛ سپس سقف اعتبار اعلام می‌شود.
#آموزش #گرفتن_اعتبار #نصب_باروک"""


def write_transcript(sid: str, title: str, duration: str, url: str, caption: str, speech_note: str, screen_note: str):
    body = f"""# رونوشت منبع {sid}

- **عنوان:** {title}
- **نشانی:** {url}
- **مدت اعلامی ویجت:** {duration or "—"}
- **مشاهده:** {OBS}
- **وضعیت دسترسی رسانه:** {ACCESS_BARRIER}

## گفتار / کپشن رسمی (متن منتشرشده روی پست)

{caption.strip()}

## آنچه روی صفحه دیده می‌شود

{screen_note}

## محدودیت شاهد

{speech_note}
"""
    (TRANS / f"{sid}.md").write_text(body, encoding="utf-8")


def make_step_card(path: Path, title: str, lines: list[str]):
    try:
        from PIL import Image, ImageDraw, ImageFont
    except ImportError:
        return
    w, h = 900, 160 + 42 * len(lines)
    img = Image.new("RGB", (w, h), (248, 246, 242))
    d = ImageDraw.Draw(img)
    d.rectangle([0, 0, w, 64], fill=(20, 70, 110))
    # default font (no Persian shaping guarantee) — use as schematic only
    d.text((24, 20), title[:80], fill=(255, 255, 255))
    y = 88
    for line in lines:
        d.text((24, y), line[:110], fill=(30, 30, 30))
        y += 42
    d.text((24, h - 36), "schematic from official caption — not a video frame", fill=(120, 120, 120))
    img.save(path)


def write_sources_csv(path: Path):
    fields = ["id", "title", "role", "product", "url", "publisher", "published", "duration", "access", "transcript", "flows"]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for s in SOURCES:
            w.writerow(s)


def build_report() -> str:
    shots = {
        "fmcg": "screenshots/official_fmcg.png",
        "home": "screenshots/official_home.png",
        "app": "screenshots/official_app.png",
        "f01": "screenshots/F01_S_schematic.png",
        "f02": "screenshots/F02_S_schematic.png",
        "f03": "screenshots/F03_S_schematic.png",
        "f07": "screenshots/F07_S_schematic.png",
        "f06": "screenshots/F06_S_schematic.png",
    }
    md = []
    md.append("# استخراج فلوهای آموزشی باروک برای سوپرمارکت و شرکت")
    md.append("")
    md.append(f"**تاریخ مشاهده:** {OBS}  ")
    md.append("**منبع اصلی کانال:** [@barook_ir](https://t.me/barook_ir) · ایندکس: [پست ۵۵](https://t.me/barook_ir/55) · جست‌وجوی [#آموزش](https://telegram.me/s/barook_ir?q=%23آموزش)")
    md.append("")
    md.append("## خلاصه اجرایی")
    md.append("")
    md.append(
        "مسیر عملیاتی غالب سوپرمارکت: ثبت‌نام/شعبه/کارت → سفارش با پخش طرف‌قرارداد → "
        "**یا** فروش اعتباری روی پوز، **یا** پیش‌فاکتور پرداخت‌نشده روی پوز هنگام تحویل، **یا** پرداخت آنلاین پس از ثبت فاکتور توسط پخش. "
        "بایت ویدئوهای تلگرام روی ماشین پژوهشی مسدود بود؛ فلوها از کپشن رسمی زمان‌دار/گام‌دار + صفحهٔ `barook.ir/fmcg` ساخته شدند. "
        "فریم جعلی از ویدئو ثبت نشده است."
    )
    md.append("")
    md.append("| شاخص | مقدار |")
    md.append("| --- | --- |")
    md.append("| فلو کامل (گام‌ها از کپشن/سایت قابل اجرا، بدون نیاز به تماشای فیلم) | ۴ (F02, F03, F06, F07) |")
    md.append("| فلو ناقص (آغاز/نتیجه یا برچسب دکمهٔ میانی کم است) | ۵ (F01, F04, F05, F08, F11) |")
    md.append("| فاقد شاهد اجرایی کافی در کپشن | چند عنوان ایندکس (توکن حقوقی، باطل کارت، پوشاک، دیجیتال) |")
    md.append("| L4 معامله واقعی | ۰ |")
    md.append("")
    md.append("## فهرست کلیک‌خور")
    md.append("")
    for item in [
        ("#مانع-دسترسی-رسانه", "مانع دسترسی رسانه"),
        ("#f01-پرداخت-آنلاین-خریدار-پست-۱۱۰", "F01 پرداخت آنلاین خریدار (۱۱۰)"),
        ("#f02-پوز-پرداخت-نشدهها-پست-۱۲۲", "F02 پوز پرداخت‌نشده‌ها (۱۲۲)"),
        ("#f03-پوز-فروش-اعتباری-پست-۱۲۱", "F03 پوز فروش اعتباری (۱۲۱)"),
        ("#f04-ثبت-سفارش-توسط-پخش-پست-۱۰۸", "F04 ثبت سفارش توسط پخش (۱۰۸)"),
        ("#f05-آپلود-فاکتور-تأمینکننده-پست-۱۴۱", "F05 آپلود فاکتور (۱۴۱)"),
        ("#f06-ثبتنام-و-فعالسازی-اعتبار-فروشگاه", "F06 ثبت‌نام و اعتبار فروشگاه"),
        ("#f07-شارژ-کیف-پول-پست-۱۵۰", "F07 شارژ کیف پول (۱۵۰)"),
        ("#f08f11-سایر-فلوهای-ناقص", "F08–F11 ناقص"),
        ("#انتقال-مسئولیت-بین-نقشها", "انتقال مسئولیت بین نقش‌ها"),
        ("#ماتریس-پوشش-نقش--فرایند", "ماتریس پوشش"),
        ("#سه-خلأ-اصلی", "سه خلأ اصلی"),
    ]:
        md.append(f"- [{item[1]}]({item[0]})")
    md.append("")
    md.append("## مانع دسترسی رسانه")
    md.append("")
    md.append(ACCESS_BARRIER)
    md.append("")
    md.append(
        "آخرین شاهد دقیق برای پست ۱۱۰: ویجت تلگرام با مدت **۰:۵۲** و کپشن کامل «پرداخت آنلاین توسط خریدار…» "
        "از `https://t.me/barook_ir/110?embed=1&mode=tme` در WebFetch. "
        "جایگزین رسمی آپارات کانال `aparat.com/barook` در این مشاهده محتوای آموزشی معادل نشان نداد. "
        "کانال بله (`ble.ir/s/barookapp`) ویدئوهای بازاریابی داشت، نه معادل ۱۱۰/۱۲۲/۱۴۱."
    )
    md.append("")
    md.append("![صفحه رسمی FMCG](" + shots["fmcg"] + ")")
    md.append("")
    md.append("*شاهد صفحهٔ عمومی — نه فریم ویدئو.*")
    md.append("")

    # F01
    md.append("## F01 — پرداخت آنلاین خریدار (پست ۱۱۰)")
    md.append("")
    md.append("| فیلد | مقدار |")
    md.append("| --- | --- |")
    md.append("| شناسه | F01 |")
    md.append("| عنوان | پرداخت اعتباری آنلاین پس از ثبت سفارش پخش |")
    md.append("| نقش | سوپرمارکت (خریدار) |")
    md.append("| آغاز | سفارش/فاکتور توسط شرکت پخش در پنل باروک ثبت شده (پیش‌نیاز: F04 / SRC-TG-108) |")
    md.append("| پیش‌نیاز | باروک‌کارت فعال؛ سفارش ثبت‌شده توسط پخش |")
    md.append("| پایان قابل مشاهده در شاهد متنی | تکمیل پرداخت آنلاین بدون حضور فیزیکی — جزئیات صفحهٔ موفقیت در کپشن نیست |")
    md.append("| وضعیت | **ناقص** — ویدئو مسدود؛ کپشن مراحل دکمه را نام نمی‌برد |")
    md.append("")
    md.append("| گام | نقش و صفحه | اقدام و متن دقیق گزینه | دادهٔ لازم | نتیجهٔ دیده‌شده و اقدام بعدی | شاهد |")
    md.append("| --- | --- | --- | --- | --- | --- |")
    md.append("| F01-S00 | شرکت پخش / پنل | ثبت سفارش فروشگاه در پنل (آموزش جدا) | شناسه فروشگاه، اقلام | سفارش آمادهٔ تأیید/پرداخت خریدار | SRC-TG-108 · 0:35 |")
    md.append("| F01-S01 | خریدار / اپ یا پنل کاربری باروک | ورود به مسیر پرداخت آنلاین همان سفارش (جزئیات دکمه فقط در ویدئو) | حساب کاربری باروک | شروع جریان پرداخت اعتباری | SRC-TG-110 · 0:52 · کپشن |")
    md.append("| F01-S02 | خریدار | طی مراحل نشان‌داده‌شده در ویدئو | — | پرداخت اعتباری آنلاین | SRC-TG-110 · حد شاهد: UI نامشخص بدون فیلم |")
    md.append("| F01-S03 | سیستم/پخش | نتیجه به شرکت برمی‌گردد (ادعا در کپشن ۱۰۸) | — | تکمیل خرید/فروش اعتباری از راه دور | SRC-TG-108+110 · اثر روی مانده در شاهد نیست |")
    md.append("")
    md.append(f"![شماتیک F01]({shots['f01']})")
    md.append("")
    md.append("**حد شاهد:** آغاز کار شرکت در فیلم ۱۱۰ نیست → به F04 ارجاع. پیام موفقیت، اثر مانده، مجوز ارسال در کپشن ۱۱۰ دیده نشد.")
    md.append("")
    md.append("مسیر موازی اعلام‌شده برای بنک‌مارکت: ۱) پرداخت آنلاین باروک پس از ثبت فاکتور در باروک ۲) کارت هنگام تحویل.")
    md.append("")

    # F02
    md.append("## F02 — پوز: پرداخت‌نشده‌ها / پیش‌فاکتور (پست ۱۲۲)")
    md.append("")
    md.append("| فیلد | مقدار |")
    md.append("| --- | --- |")
    md.append("| شناسه | F02 |")
    md.append("| نقش | موزع/پخش روی پوز + دارنده کارت |")
    md.append("| آغاز | اپ باروک روی POS باز است؛ پیش‌فاکتور صادر شده |")
    md.append("| پایان | کشیدن کارت توسط خریدار |")
    md.append("| وضعیت | **کامل در سطح کپشن اجرایی** (بدون فریم) |")
    md.append("")
    md.append("| گام | نقش و صفحه | اقدام و متن دقیق گزینه | دادهٔ لازم | نتیجه و بعدی | شاهد |")
    md.append("| --- | --- | --- | --- | --- | --- |")
    md.append("| F02-S01 | موزع / اپ باروک روی POS | ورود به اپلیکیشن باروک | — | صفحهٔ اصلی اپ پوز | SRC-TG-122 |")
    md.append("| F02-S02 | موزع / سفارشات | باز کردن بخش «سفارشات» | — | فهرست سفارشات | SRC-TG-122 |")
    md.append("| F02-S03 | موزع | کلیک روی «پرداخت نشده‌ها» | — | لیست پیش‌فاکتورهای باز | SRC-TG-122 · متن دقیق گزینه |")
    md.append("| F02-S04 | موزع / جزئیات | کلیک روی «پرداخت» | انتخاب همان پیش‌فاکتور | درخواست کشیدن کارت | SRC-TG-122 |")
    md.append("| F02-S05 | خریدار | کشیدن باروک‌کارت روی دستگاه | کارت فیزیکی | تکمیل پرداخت از دید کپشن | SRC-TG-122 · رمز در این کپشن ذکر نشده (در F03 هست) |")
    md.append("")
    md.append(f"![شماتیک F02]({shots['f02']})")
    md.append("")
    md.append("**نکتهٔ تحویل:** مأموریت «رویهٔ تحویل» را به ۱۲۲ نگاشته؛ شاهدِ متنی همان تسویهٔ پیش‌فاکتور روی پوز هنگام حضور کارت است، نه چک‌لیست انبار/رسید کالا.")
    md.append("")

    # F03
    md.append("## F03 — پوز: فروش اعتباری (پست ۱۲۱)")
    md.append("")
    md.append("| فیلد | مقدار |")
    md.append("| --- | --- |")
    md.append("| شناسه | F03 |")
    md.append("| نقش | اپراتور پوز + خریدار |")
    md.append("| آغاز | اپ باروک روی POS |")
    md.append("| پایان | کشیدن کارت + ورود رمز |")
    md.append("| وضعیت | **کامل در سطح کپشن اجرایی** |")
    md.append("")
    md.append("| گام | نقش و صفحه | اقدام و متن دقیق گزینه | دادهٔ لازم | نتیجه و بعدی | شاهد |")
    md.append("| --- | --- | --- | --- | --- | --- |")
    md.append("| F03-S01 | اپراتور / POS | باز کردن اپلیکیشن باروک | — | منوی باروک | SRC-TG-121 |")
    md.append("| F03-S02 | اپراتور | کلیک «فروش اعتباری» | — | فرم مبلغ/سررسید | SRC-TG-121 · متن گزینه |")
    md.append("| F03-S03 | اپراتور | وارد کردن مبلغ جمع فاکتور و مدت سررسید (از نماینده پخش بپرس) | مبلغ؛ سررسید | آمادهٔ تأیید | SRC-TG-121 |")
    md.append("| F03-S04 | اپراتور | کلیک دکمه «فروش اعتباری» | — | درخواست کارت | SRC-TG-121 |")
    md.append("| F03-S05 | خریدار | کشیدن باروک‌کارت و ورود رمز | کارت + رمز | تراکنش اعتباری از دید کپشن | SRC-TG-121 |")
    md.append("")
    md.append(f"![شماتیک F03]({shots['f03']})")
    md.append("")
    md.append("نسخهٔ قدیمی‌تر با کپشن نازک: SRC-TG-017 (پست ۱۷). اختلاف نقش/تاریخ: ۱۷ قدیمی‌تر از ۱۲۱/۱۲۲.")
    md.append("")

    # F04
    md.append("## F04 — ثبت سفارش توسط پخش برای پرداخت/تأیید از راه دور (پست ۱۰۸)")
    md.append("")
    md.append("| وضعیت | **ناقص** — کپشن هدف را می‌گوید؛ دکمه‌های پنل را نام نمی‌برد |")
    md.append("")
    md.append("| گام | نقش و صفحه | اقدام | داده | نتیجه | شاهد |")
    md.append("| --- | --- | --- | --- | --- | --- |")
    md.append("| F04-S01 | پخش / پنل باروک | ثبت سفارش فروشگاه‌دار به روش نمایش‌داده‌شده در ویدئو | شناسه فروشگاه؛ اقلام | سفارش ثبت می‌شود | SRC-TG-108 · 0:35 |")
    md.append("| F04-S02 | خریدار | تأیید خرید اعتباری / پرداخت (F01) | — | تکمیل دور از محل | SRC-TG-108+110 |")
    md.append("")

    # F05
    md.append("## F05 — آپلود فاکتور تأمین‌کننده (پست ۱۴۱)")
    md.append("")
    md.append("| وضعیت | **ناقص** — فقط وعدهٔ قدم‌به‌قدم در کپشن؛ نام منو/فیلد نیست |")
    md.append("")
    md.append("| گام | نقش | اقدام | شاهد |")
    md.append("| --- | --- | --- | --- |")
    md.append("| F05-S01 | تأمین‌کننده / پنل | ورود به پنل تأمین‌کننده باروک | SRC-TG-141 · 1:41 |")
    md.append("| F05-S02 | تأمین‌کننده | آپلود فاکتورها (جزئیات UI فقط در ویدئو) | SRC-TG-141 · حد شاهد |")
    md.append("")

    # F06
    md.append("## F06 — ثبت‌نام و فعال‌سازی اعتبار فروشگاه")
    md.append("")
    md.append("| شناسه | F06 |")
    md.append("| نقش | سوپرمارکت |")
    md.append("| منابع | کپشن آموزشی ثبت‌نام در فهرست #آموزش + `barook.ir/fmcg` + راهنمای شریک عظیمی |")
    md.append("| وضعیت | **کامل در سطح سند عمومی** (بدون نیاز به فیلم برای مسیر اصلی) |")
    md.append("")
    md.append("| گام | نقش و صفحه | اقدام و متن گزینه | داده | نتیجه و بعدی | شاهد |")
    md.append("| --- | --- | --- | --- | --- | --- |")
    md.append("| F06-S01 | خریدار / barook.ir | ثبت‌نام / درخواست اعتبار | — | شروع پرونده | SRC-OFF-FMCG · کپشن #آموزش |")
    md.append("| F06-S02 | خریدار | بارگذاری هویت، فروشگاه، کارت ملی، سند کسب | مدارک | پرونده ناقص→کامل | SRC-OFF-FMCG |")
    md.append("| F06-S03 | خریدار | کلیک «مراحل دریافت اعتبار را تکمیل کنید» | — | انتظار چند ساعت | کپشن ثبت‌نام آموزشی |")
    md.append("| F06-S04 | خریدار | مشاهده نتیجه در همان بخش؛ کلیک «ادامه مراحل بانکی» | — | انتخاب شعبه | کپشن ثبت‌نام |")
    md.append("| F06-S05 | خریدار | انتخاب شعبه تجارت؛ بستن نوبت با ✖️؛ دیدن تاریخ مراجعه در صفحه اصلی | شعبه | نوبت بانک | کپشن ثبت‌نام |")
    md.append("| F06-S06 | خریدار+بانک / شعبه | مراجعه با مدارک؛ چک ضمانت ۱۲۵٪ برای A/B | چک؛ مدارک | سقف اعتبار؛ ارسال کارت | SRC-OFF-FMCG |")
    md.append("| F06-S07 | خریدار | اعلام داشتن کارت به ویزیتور؛ مذاکره سررسید | — | آمادهٔ خرید | SRC-OFF-FMCG |")
    md.append("| F06-S08 | خریدار | کشیدن کارت روی کارتخوان باروک مأمور تحویل | کارت | مصرف اعتبار | SRC-OFF-FMCG |")
    md.append("| F06-S09 | خریدار / حساب جاری تجارت | تأمین وجه قبل از سررسید | وجه | کسر اصل+۲٫۲۵٪؛ برگشت سقف | SRC-OFF-FMCG |")
    md.append("")
    md.append(f"![شماتیک F06]({shots['f06']})")
    md.append("")

    # F07
    md.append("## F07 — شارژ کیف پول (پست ۱۵۰)")
    md.append("")
    md.append("| وضعیت | **کامل در سطح کپشن اجرایی** |")
    md.append("")
    md.append("| گام | نقش و صفحه | اقدام | داده | نتیجه | شاهد |")
    md.append("| --- | --- | --- | --- | --- | --- |")
    md.append("| F07-S01 | خریدار / اپ | صفحه اصلی اپ باروک | — | — | SRC-TG-150 |")
    md.append("| F07-S02 | خریدار | دکمه «کیف پول» پایین اپ → «شارژ کیف پول» | — | صفحه مبلغ | SRC-TG-150 · متن گزینه |")
    md.append("| F07-S03 | خریدار | مبلغ پیش‌فرض یا مبلغ دلخواه → درگاه پرداخت | مبلغ | پرداخت | SRC-TG-150 |")
    md.append("| F07-S04 | سیستم | اعلام موفقیت شارژ در کپشن | — | آمادهٔ خرید با کیف | SRC-TG-150 · اثر حساب بانکی خارج از شاهد |")
    md.append("")
    md.append(f"![شماتیک F07]({shots['f07']})")
    md.append("")

    # F08-F11
    md.append("## F08–F11 سایر فلوهای ناقص")
    md.append("")
    md.append("| فلو | منبع | آنچه داریم | آنچه کم است |")
    md.append("| --- | --- | --- | --- |")
    md.append("| F08 فعالسازی کارت | SRC-TG-013 · 1:08 | هدف آموزش | گام UI |")
    md.append("| F09 نصب روی P10 | SRC-TG-058 · 3:25 | مدل دستگاه توسن‌تکنو P10؛ هدف نصب | گام نصب |")
    md.append("| F10 به‌روزرسانی اپ پوز | SRC-TG-060 · 1:07 | هدف به‌روزرسانی | منو/دکمه |")
    md.append("| F11 کسب‌نو حقوقی | SRC-TG-109 · 0:58 | مدیریت حق امضا مطابق روزنامه رسمی؛ مدیریت فاکتور | مسیر با/بدون توکن در این کپشن جدا نشده؛ ایندکس ۵۵ به آموزش‌های جدا اشاره دارد |")
    md.append("")
    md.append("عناوین ایندکس ۵۵ که در این دور شناسهٔ پستشان به کپشن گام‌دار نرسید (شاهد پیدا نشد / فقط عنوان): ثبت‌نام کسب‌نو دیجیتال، پوشاک، حقوقی بدون توکن، حقوقی با توکن، باطل کردن کارت، لیست پخش استانی.")
    md.append("")

    md.append("## انتقال مسئولیت بین نقش‌ها")
    md.append("")
    md.append("| نقطه | از | به | شاهد |")
    md.append("| --- | --- | --- | --- |")
    md.append("| ثبت سفارش/فاکتور در پنل | پخش | خریدار می‌تواند آنلاین بپردازد | SRC-TG-108 → SRC-TG-110 |")
    md.append("| انتخاب پرداخت‌نشده + پرداخت | موزع روی پوز | خریدار کارت می‌کشد | SRC-TG-122 |")
    md.append("| ورود مبلغ/سررسید فروش اعتباری | اپراتور پوز | خریدار کارت+رمز | SRC-TG-121 |")
    md.append("| نتیجه پرداخت از راه دور | خریدار | تکمیل معامله برای پخش (ادعا) | SRC-TG-108 کپشن |")
    md.append("")

    md.append("## ماتریس پوشش نقش × فرایند")
    md.append("")
    md.append("| فرایند | سوپرمارکت | حقوقی متقاضی | پخش/تأمین‌کننده |")
    md.append("| --- | --- | --- | --- |")
    md.append("| ثبت‌نام و احراز | F06 کامل‌متن | F11 ناقص | — |")
    md.append("| مدارک فروشگاه | F06 | عنوان ایندکس؛ جزئیات کم | — |")
    md.append("| درخواست/فعالسازی اعتبار | F06 | F11 + عناوین توکن | — |")
    md.append("| مشاهده مانده/سقف | صفحه FMCG ادعا؛ UI داخل حساب نه | — | — |")
    md.append("| دریافت/فعالسازی کارت | F08 ناقص + F06 ارسال کارت | — | — |")
    md.append("| یافتن سفارش/فاکتور | F01 ناقص UI | — | F04/F05 |")
    md.append("| پرداخت اعتباری آنلاین | F01 ناقص UI | — | پیش‌نیاز F04 |")
    md.append("| پرداخت با کارت‌خوان | F02/F03 | — | F02/F03 |")
    md.append("| پیگیری موفقیت/خطا | شاهد پیدا نشد | شاهد پیدا نشد | شاهد پیدا نشد |")
    md.append("| بازپرداخت | F06-S09 سند رسمی | — | تسویه از بانک (ادعا) |")
    md.append("| تحویل | مسیر کارت هنگام تحویل (FMCG+بنک) | — | F02 مرتبط پوز |")
    md.append("| مغایرت/مرجوعی | شاهد پیدا نشد | — | شاهد پیدا نشد |")
    md.append("| پذیرندگی/نصب پوز | — | — | F09/F10 ناقص UI |")
    md.append("| آپلود فاکتور | — | — | F05 ناقص UI |")
    md.append("| شارژ کیف | F07 کامل‌کپشن | — | — |")
    md.append("")

    md.append("## سه خلأ اصلی")
    md.append("")
    md.append("1. **بایت ویدئو تلگرام ۱۱۰/۱۲۲/۱۴۱ و بقیه** — نوع شاهد لازم: فایل mp4 رسمی یا میزبانی آپارات معادل + رونوشت زمان‌دار از خود رسانه (بستن G-BK-06).")
    md.append("2. **صفحهٔ موفقیت پرداخت آنلاین و اثر مانده** — نوع شاهد: اسکرین ماسک‌شدهٔ اپ خریدار پس از F01 یا سند تراکنش بی‌نام.")
    md.append("3. **آموزش‌های حقوقی با/بدون توکن و باطل کارت** — نوع شاهد: پست‌های لینک‌شده از ایندکس ۵۵ با کپشن/ویدئو گام‌دار.")
    md.append("")
    md.append("## پیوست: اتصال به پژوهش قبلی")
    md.append("")
    md.append("- G-BK-06 همچنان باز است از حیث فریم/رونوشت زمان‌دار ویدئو؛ این بسته کپشن‌ها و فلوهای متنی را جلو می‌برد.")
    md.append("- OrderSteps قبلی (OS-BK-*) با F02/F03/F06 هم‌راستا است؛ F01 مسیر موازی آنلاین را اضافه می‌کند و نباید با پنل فرضی سناریوی ۰٫۱۰ یکی گرفته شود.")
    md.append("")
    return "\n".join(md)


def main():
    # transcripts
    for sid, cap in CAPTIONS.items():
        src = next(s for s in SOURCES if s["id"] == sid)
        write_transcript(
            sid,
            src["title"],
            src.get("duration") or "",
            src["url"],
            cap,
            speech_note="گفتار داخل ویدئو شنیده/پیاده‌سازی نشد چون فایل رسانه دریافت نشد. متن بالا کپشن رسمی پست است.",
            screen_note="فریم‌های UI از ویدئو استخراج نشد (مانع دسترسی). برای منابع با کپشن گام‌دار، همان برچسب گزینه‌ها به‌عنوان شاهد متنی ثبت شده‌اند.",
        )
    # registration caption as extra transcript
    write_transcript(
        "SRC-TG-REG",
        "آموزش گرفتن اعتبار / ثبت‌نام (کپشن فهرست #آموزش)",
        "3:04",
        "https://telegram.me/s/barook_ir?q=%23آموزش",
        REG_CAPTION_FROM_SEARCH,
        "شناسهٔ دقیق پست در این دور از روی جست‌وجو به embed پایدار وصل نشد؛ متن از فهرست #آموزش است.",
        "مراحل دکمه در کپشن آمده؛ فریم ویدئو نیست.",
    )

    # schematics
    make_step_card(
        SHOT / "F01_S_schematic.png",
        "F01 online pay (caption-level)",
        [
            "S00 Distributor registers order in Barook panel (post 108)",
            "S01 Buyer opens Barook user flow for that invoice (post 110)",
            "S02 Buyer completes on-screen steps shown in video [UI unknown]",
            "S03 Remote credit purchase completes (claimed)",
        ],
    )
    make_step_card(
        SHOT / "F02_S_schematic.png",
        "F02 POS unpaid (post 122 caption)",
        [
            "Open Barook app on POS",
            "Orders -> Unpaid (پرداخت نشده‌ها)",
            "Details -> Pay (پرداخت)",
            "Buyer swipes BarookCard",
        ],
    )
    make_step_card(
        SHOT / "F03_S_schematic.png",
        "F03 POS credit sale (post 121 caption)",
        [
            "Open Barook on POS -> Credit sale (فروش اعتباری)",
            "Enter invoice total + maturity",
            "Tap Credit sale",
            "Swipe card + enter PIN",
        ],
    )
    make_step_card(
        SHOT / "F07_S_schematic.png",
        "F07 wallet top-up (post 150 caption)",
        [
            "App home",
            "Wallet (bottom) -> Charge wallet",
            "Preset or custom amount -> gateway",
            "Wallet charged (per caption)",
        ],
    )
    make_step_card(
        SHOT / "F06_S_schematic.png",
        "F06 onboarding (fmcg + training caption)",
        [
            "Register + upload ID/store docs",
            "Complete credit steps / bank branch",
            "Branch visit + 125% cheque (A/B)",
            "Receive card -> swipe on distributor POS",
            "Repay before due -> ceiling refreshes",
        ],
    )

    report = build_report()
    (ROOT / "Barook-Training-Flows-FA.md").write_text(report, encoding="utf-8")
    (ART / "Barook-Training-Flows-FA.md").write_text(report, encoding="utf-8")

    write_sources_csv(ROOT / "sources.csv")
    write_sources_csv(ART / "sources.csv")

    # copy trees
    for sub in ("transcripts", "screenshots"):
        dst = ART / sub
        if dst.exists():
            shutil.rmtree(dst)
        shutil.copytree(ROOT / sub, dst)

    meta = {
        "observed_at": OBS,
        "flows_complete_caption_level": ["F02", "F03", "F06", "F07"],
        "flows_partial": ["F01", "F04", "F05", "F08", "F09", "F10", "F11"],
        "video_bytes": 0,
        "access_barrier": ACCESS_BARRIER,
        "primary_post_110": "https://t.me/barook_ir/110",
    }
    (ROOT / "build_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    (ART / "build_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(meta, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
