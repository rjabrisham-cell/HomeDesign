# -*- coding: utf-8 -*-
import re, json, ssl, urllib.request
from pathlib import Path

OUT = Path(__file__).resolve().parent / "raw"
OUT.mkdir(parents=True, exist_ok=True)
ctx = ssl.create_default_context()
HDR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

POSTS = [110, 122, 141] + list(range(100, 200))
# also search page
URLS = [
    ("search_amoozesh", "https://t.me/s/barook_ir?q=%23%D8%A2%D9%85%D9%88%D8%B2%D8%B4"),
    ("s_110", "https://t.me/s/barook_ir/110"),
    ("s_122", "https://t.me/s/barook_ir/122"),
    ("s_141", "https://t.me/s/barook_ir/141"),
]

def get(url):
    req = urllib.request.Request(url, headers=HDR)
    return urllib.request.urlopen(req, context=ctx, timeout=45).read().decode("utf-8", "replace")

def parse_messages(html):
    msgs = []
    # split by message widgets
    parts = re.split(r'(?=<div class="tgme_widget_message )', html)
    for p in parts:
        mpost = re.search(r'data-post="barook_ir/(\d+)"', p)
        if not mpost:
            continue
        pid = int(mpost.group(1))
        caption = ""
        cm = re.search(r'class="tgme_widget_message_text[^"]*"[^>]*>(.*?)</div>', p, re.S)
        if cm:
            caption = re.sub(r"<br\s*/?>", "\n", cm.group(1))
            caption = re.sub(r"<[^>]+>", "", caption)
            caption = re.sub(r"&nbsp;", " ", caption)
            caption = re.sub(r"&quot;", '"', caption)
            caption = re.sub(r"&amp;", "&", caption).strip()
        dur = None
        dm = re.search(r'tgme_widget_message_video_duration[^>]*>([^<]+)', p)
        if dm:
            dur = dm.group(1).strip()
        vids = re.findall(r'video_player[^>]+src="(https://cdn[^"]+)"', p)
        if not vids:
            vids = re.findall(r'(https://cdn\d+\.telesco\.pe/file/[^"\']+\.mp4[^"\']*)', p)
        thumbs = re.findall(r'(https://cdn\d+\.telesco\.pe/file/[^"\']+\.jpg[^"\']*)', p)
        date = None
        dt = re.search(r'datetime="([^"]+)"', p)
        if dt:
            date = dt.group(1)
        msgs.append({
            "id": pid,
            "caption": caption,
            "duration": dur,
            "videos": vids,
            "thumbs": thumbs[:3],
            "date": date,
            "hashtag_amoozesh": "#آموزش" in caption or "آموزش" in caption[:80],
        })
    return msgs

all_msgs = {}
for name, url in URLS:
    try:
        html = get(url)
        (OUT / f"{name}.html").write_text(html, encoding="utf-8")
        print(f"OK {name} {len(html)}")
        for m in parse_messages(html):
            all_msgs[m["id"]] = m
    except Exception as e:
        print(f"ERR {name}: {e}")

# fetch individual embeds for key posts and neighbors that look like trainings
for pid in sorted(set([110, 122, 141] + list(all_msgs.keys()))):
    if pid in all_msgs and all_msgs[pid].get("videos"):
        continue
    url = f"https://t.me/barook_ir/{pid}?embed=1&mode=tme"
    try:
        html = get(url)
        (OUT / f"embed_{pid}.html").write_text(html, encoding="utf-8")
        for m in parse_messages(html):
            prev = all_msgs.get(m["id"], {})
            # merge preferring non-empty
            merged = {**prev, **{k: v for k, v in m.items() if v}}
            if prev.get("caption") and not m.get("caption"):
                merged["caption"] = prev["caption"]
            all_msgs[m["id"]] = merged
        print(f"embed {pid} ok")
    except Exception as e:
        print(f"embed ERR {pid}: {e}")

# also crawl s/ pages around key posts for more context
for pid in [105, 108, 109, 110, 111, 112, 120, 121, 122, 123, 140, 141, 142, 150, 160, 170, 180]:
    url = f"https://t.me/s/barook_ir/{pid}"
    try:
        html = get(url)
        for m in parse_messages(html):
            if m["id"] not in all_msgs or (m.get("videos") and not all_msgs[m["id"]].get("videos")):
                all_msgs[m["id"]] = m
            elif m.get("caption") and len(m["caption"]) > len(all_msgs[m["id"]].get("caption") or ""):
                all_msgs[m["id"]]["caption"] = m["caption"]
                if m.get("videos"):
                    all_msgs[m["id"]]["videos"] = m["videos"]
                if m.get("duration"):
                    all_msgs[m["id"]]["duration"] = m["duration"]
        print(f"s/{pid} msgs={len(parse_messages(html))}")
    except Exception as e:
        print(f"s ERR {pid}: {e}")

(OUT / "messages.json").write_text(
    json.dumps([all_msgs[k] for k in sorted(all_msgs)], ensure_ascii=False, indent=2),
    encoding="utf-8",
)
print("TOTAL", len(all_msgs))
# print training-like
for m in sorted(all_msgs.values(), key=lambda x: x["id"]):
    cap = (m.get("caption") or "").replace("\n", " ")[:120]
    if m["id"] in (110, 122, 141) or "آموزش" in (m.get("caption") or "") or m.get("duration"):
        print(f"#{m['id']} dur={m.get('duration')} vids={len(m.get('videos') or [])} | {cap}")
