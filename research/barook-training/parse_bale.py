# -*- coding: utf-8 -*-
import re, json, ssl, urllib.request
from pathlib import Path
from html import unescape

OUT = Path(__file__).resolve().parent / "raw"
VID = Path(__file__).resolve().parent / "videos"
SHOT = Path(__file__).resolve().parent / "screenshots"
VID.mkdir(exist_ok=True)
SHOT.mkdir(exist_ok=True)
ctx = ssl.create_default_context()
HDR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

html = (OUT / "bale_s_barookapp.html").read_text(encoding="utf-8")
print("html len", len(html))

# Find JSON blobs
# Often Next/SSR embeds __NEXT_DATA__ or similar
for marker in ["__NEXT_DATA__", "window.__INITIAL", "posts", "documentExVideo", "proxy_download"]:
    print(marker, html.find(marker))

# Extract all proxy_download URLs
proxies = sorted(set(re.findall(r'https://file-gw\d*\.ble\.ir/proxy_download/[^"\\]+', html)))
print("proxy urls", len(proxies))
for p in proxies[:20]:
    print(" ", p[:120])

# Extract video metadata blocks
videos = []
for m in re.finditer(r'documentExVideo":\{([^}]+)\}', html):
    block = m.group(0)
    videos.append(block)
print("video blocks", len(videos))
for v in videos[:15]:
    print(" ", v[:200])

# Extract message texts near آموزش
# unescape unicode escapes in JSON strings
def decode_js_str(s):
    try:
        return bytes(s, "utf-8").decode("unicode_escape")
    except Exception:
        return s

# Find caption-like strings
caps = re.findall(r'"text"\s*:\s*"((?:[^"\\]|\\.){20,800})"', html)
print("text fields", len(caps))
amoozesh_posts = []
for c in caps:
    t = decode_js_str(c).replace("\\n", "\n")
    if any(k in t for k in ["آموزش", "پوز", "فاکتور", "کیف پول", "فعالسازی", "فعال‌سازی", "پرداخت", "توکن", "ثبت نام", "ثبت‌نام"]):
        amoozesh_posts.append(t)
        print("---CAP---")
        print(t[:400])

# Also search for message objects with id
msgs = []
for m in re.finditer(r'\{"id":(\d+),"date":(\d+),.*?"text":"((?:[^"\\]|\\.)*?)"', html):
    mid, date, text = m.group(1), m.group(2), decode_js_str(m.group(3)).replace("\\n", "\n")
    msgs.append({"id": int(mid), "date": int(date), "text": text})

print("parsed msgs", len(msgs))
(OUT / "bale_msgs.json").write_text(json.dumps(msgs, ensure_ascii=False, indent=2), encoding="utf-8")
(OUT / "bale_amoozesh_caps.json").write_text(json.dumps(amoozesh_posts, ensure_ascii=False, indent=2), encoding="utf-8")

# Try download first few proxy urls to see content-type
for i, url in enumerate(proxies[:5]):
    try:
        req = urllib.request.Request(unescape(url), headers=HDR)
        with urllib.request.urlopen(req, context=ctx, timeout=60) as r:
            data = r.read(64)
            ctype = r.headers.get("Content-Type", "")
            print(f"DL{i} ctype={ctype} magic={data[:12]} len_hdr={r.headers.get('Content-Length')}")
    except Exception as e:
        print(f"DL{i} ERR", e)
