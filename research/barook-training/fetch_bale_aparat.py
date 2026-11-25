# -*- coding: utf-8 -*-
import re, json, ssl, urllib.request, urllib.parse
from pathlib import Path

OUT = Path(__file__).resolve().parent / "raw"
VID = Path(__file__).resolve().parent / "videos"
OUT.mkdir(parents=True, exist_ok=True)
VID.mkdir(parents=True, exist_ok=True)
ctx = ssl.create_default_context()
HDR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"}

def get(url, binary=False):
    req = urllib.request.Request(url, headers=HDR)
    with urllib.request.urlopen(req, context=ctx, timeout=60) as r:
        data = r.read()
    return data if binary else data.decode("utf-8", "replace")

# --- Aparat ---
aparat_urls = [
    "https://www.aparat.com/barook",
    "https://www.aparat.com/result/%D8%A8%D8%A7%D8%B1%D9%88%DA%A9",
    "https://www.aparat.com/result/barook",
    "https://www.aparat.com/result/%D8%A8%D8%A7%D8%B1%D9%88%DA%A9%20%DA%A9%D8%A7%D8%B1%D8%AA",
]
aparat_hits = []
for u in aparat_urls:
    try:
        html = get(u)
        name = "aparat_" + re.sub(r"[^a-zA-Z0-9]+", "_", u.split(".com/")[-1])[:40]
        (OUT / f"{name}.html").write_text(html, encoding="utf-8")
        # titles and video links
        vids = re.findall(r'href="(/v/[^"]+)"[^>]*>([^<]{0,120})', html)
        titles = re.findall(r'"title"\s*:\s*"([^"]{5,120})"', html)
        usernames = re.findall(r'"username"\s*:\s*"([^"]+)"', html)
        print(f"APARAT {u} len={len(html)} href_vids={len(vids)} titles={len(titles)}")
        for href, t in vids[:30]:
            aparat_hits.append({"href": "https://www.aparat.com" + href, "title": t.strip(), "from": u})
            print("  v:", t.strip()[:80], href)
        for t in titles[:20]:
            if "باروک" in t or "barook" in t.lower() or "کسب" in t:
                print("  title:", t)
        print("  users:", usernames[:10])
    except Exception as e:
        print("APARAT ERR", u, e)

(OUT / "aparat_hits.json").write_text(json.dumps(aparat_hits, ensure_ascii=False, indent=2), encoding="utf-8")

# --- Bale ---
bale_urls = [
    "https://ble.ir/s/barookapp",
    "https://ble.ir/barookapp",
]
for u in bale_urls:
    try:
        html = get(u)
        name = "bale_" + re.sub(r"[^a-zA-Z0-9]+", "_", u.split("ble.ir/")[-1])[:40]
        (OUT / f"{name}.html").write_text(html, encoding="utf-8")
        print(f"BALE {u} len={len(html)}")
        # look for video / media
        for pat in [r'https?://[^"\']+\.mp4[^"\']*', r'https?://cdn[^"\']+', r'data-src="([^"]+)"', r'src="(https://[^"]+)"']:
            found = re.findall(pat, html)
            print(f"  pat {pat[:30]} count={len(found)} sample={found[:3]}")
        # text snippets with آموزش
        texts = re.findall(r'آموزش[^<]{0,200}', html)
        print("  amoozesh snippets:", len(texts))
        for t in texts[:15]:
            print("   ", t[:150].replace("\n", " "))
    except Exception as e:
        print("BALE ERR", u, e)

# try bale API-ish endpoints sometimes used
for u in [
    "https://ble.ir/s/barookapp?q=%D8%A2%D9%85%D9%88%D8%B2%D8%B4",
]:
    try:
        html = get(u)
        (OUT / "bale_search_amoozesh.html").write_text(html, encoding="utf-8")
        print("BALE search len", len(html))
    except Exception as e:
        print("BALE search ERR", e)
