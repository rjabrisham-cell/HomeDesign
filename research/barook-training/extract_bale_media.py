# -*- coding: utf-8 -*-
"""Extract Bale channel posts + media URLs and download training videos."""
import re, json, ssl, urllib.request
from pathlib import Path
from html import unescape

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "raw"
VID = ROOT / "videos"
SHOT = ROOT / "screenshots"
for p in (VID, SHOT):
    p.mkdir(exist_ok=True)

ctx = ssl.create_default_context()
HDR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

html = (OUT / "bale_s_barookapp.html").read_text(encoding="utf-8")
# Also search page
search = OUT / "bale_search_amoozesh.html"
if search.exists():
    html2 = search.read_text(encoding="utf-8")
    if len(html2) > len(html):
        html = html2

# Unescape common JS unicode in the page for easier parsing
raw = html

# Find all proxy_download full URLs (may be truncated in previous regex)
proxies = re.findall(r'https://file-gw\d+\.ble\.ir/proxy_download/[^"\\\s<>]+', raw)
# Also try truncated with =?u=
proxies2 = re.findall(r'https://file-gw\d+\.ble\.ir/proxy_download/[^"]{10,500}', raw)
all_proxies = []
seen = set()
for p in proxies + proxies2:
    p = unescape(p).rstrip('\\')
    # cut at first whitespace or quote residue
    p = re.split(r'[\\"\s]', p)[0]
    if p not in seen and "proxy_download" in p:
        seen.add(p)
        all_proxies.append(p)

print("unique proxies", len(all_proxies))

# Probe each
meta = []
for i, url in enumerate(all_proxies):
    try:
        req = urllib.request.Request(url, headers={**HDR, "Range": "bytes=0-1023"})
        with urllib.request.urlopen(req, context=ctx, timeout=45) as r:
            chunk = r.read(32)
            ctype = r.headers.get("Content-Type", "")
            clen = r.headers.get("Content-Length") or r.headers.get("Content-Range", "")
            kind = "video" if "video" in ctype or chunk[4:8] == b"ftyp" else (
                "image" if "image" in ctype or chunk[:2] == b"\xff\xd8" else "other"
            )
            meta.append({"i": i, "url": url, "ctype": ctype, "clen": clen, "kind": kind, "magic": chunk[:12].hex()})
            print(f"{i:02d} {kind:5} {ctype:20} {clen} {url[:90]}")
    except Exception as e:
        meta.append({"i": i, "url": url, "error": str(e)})
        print(f"{i:02d} ERR {e}")

(OUT / "bale_proxy_meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

# Extract surrounding context for each proxy URL (caption clues)
contexts = []
for url in all_proxies:
    idx = raw.find(url[:60])
    if idx < 0:
        continue
    window = raw[max(0, idx - 2500): idx + 200]
    # find Persian-ish text nearby
    texts = re.findall(r'[\u0600-\u06FF][\u0600-\u06FF\s\d\w\.\,\!\?\:\-\(\)«»\"\'\/]{15,300}', window)
    contexts.append({"url": url, "nearby_fa": texts[-8:], "window_snip": window[-400:]})

(OUT / "bale_proxy_context.json").write_text(json.dumps(contexts, ensure_ascii=False, indent=2), encoding="utf-8")
for c in contexts:
    print("CTX", c["url"][:70])
    for t in c["nearby_fa"][-4:]:
        print("   ", t[:120])

# Download all videos
videos = [m for m in meta if m.get("kind") == "video"]
print("videos to download", len(videos))
for j, m in enumerate(videos):
    dest = VID / f"bale_video_{j:02d}.mp4"
    if dest.exists() and dest.stat().st_size > 100000:
        print("skip existing", dest.name, dest.stat().st_size)
        continue
    print("downloading", dest.name, "...")
    try:
        req = urllib.request.Request(m["url"], headers=HDR)
        with urllib.request.urlopen(req, context=ctx, timeout=180) as r:
            data = r.read()
        dest.write_bytes(data)
        print(" saved", len(data))
    except Exception as e:
        print(" DL ERR", e)

# Download images as potential thumbs/screenshots
imgs = [m for m in meta if m.get("kind") == "image"]
for j, m in enumerate(imgs):
    dest = SHOT / f"bale_img_{j:02d}.jpg"
    try:
        req = urllib.request.Request(m["url"], headers=HDR)
        with urllib.request.urlopen(req, context=ctx, timeout=60) as r:
            dest.write_bytes(r.read())
        print("img", dest.name, dest.stat().st_size)
    except Exception as e:
        print("img ERR", e)
