# -*- coding: utf-8 -*-
"""Fetch Telegram embeds/videos via SOCKS5 or Chrome channel."""
import json, re, ssl, urllib.request
from pathlib import Path
from urllib.parse import urlparse

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "raw"
VID = ROOT / "videos"
RAW.mkdir(exist_ok=True)
VID.mkdir(exist_ok=True)

# Fetch proxy list
proxies = []
try:
    import urllib.request as u
    ctx = ssl.create_default_context()
    req = u.Request(
        "https://api.telegramvpn.org/proxies?status=active&sort=ping&type=socks5&per=20&for=IR",
        headers={"User-Agent": "Mozilla/5.0"},
    )
    data = json.loads(u.urlopen(req, context=ctx, timeout=30).read().decode())
    for p in data.get("proxies") or []:
        proxies.append(f"socks5://{p['server']}:{p['port']}")
    print("echo proxies", len(proxies), proxies[:5])
except Exception as e:
    print("echo api fail", e)

# fallback static
proxies += [
    "socks5://80.191.243.212:1080",
    "socks5://5.144.133.195:9050",
    "socks5://45.139.10.136:1080",
]

# Try PySocks
try:
    import socks
    import socket
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "PySocks", "-q"])
    import socks
    import socket

POSTS = [110, 122, 141, 108, 121, 109]

def try_fetch(proxy_url, target, out_path=None, binary=False):
    parsed = urlparse(proxy_url)
    host, port = parsed.hostname, parsed.port or 1080
    # monkeypatch
    socks.set_default_proxy(socks.SOCKS5, host, port)
    socket.socket = socks.socksocket
    try:
        req = urllib.request.Request(target, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=40) as r:
            data = r.read()
        if out_path:
            Path(out_path).write_bytes(data)
        return True, data[:200] if not binary else len(data), data
    except Exception as e:
        return False, str(e), None
    finally:
        # restore? hard; reimport socket if needed
        pass

# Restore socket between attempts by reloading
import importlib

working = None
for px in proxies[:12]:
    importlib.reload(socket)
    socks.set_default_proxy(socks.SOCKS5, urlparse(px).hostname, urlparse(px).port or 1080)
    socket.socket = socks.socksocket
    ok, info, data = False, None, None
    try:
        req = urllib.request.Request("https://t.me/barook_ir/110?embed=1&mode=tme", headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=35) as r:
            data = r.read()
        ok, info = True, len(data)
    except Exception as e:
        ok, info = False, str(e)
    print("proxy", px, "OK" if ok else "FAIL", info if not ok else f"len={info}")
    if ok:
        working = px
        (RAW / "proxy_embed_110.html").write_bytes(data)
        break

if not working:
    print("NO WORKING PROXY — try chrome channel")
else:
    html = data.decode("utf-8", "replace")
    vids = re.findall(r'(https://cdn\d+\.telesco\.pe/file/[^"\']+)', html)
    print("vids found", vids[:3])
    # also video_player src
    vids2 = re.findall(r'video[^>]+src="(https://[^"]+)"', html)
    print("vids2", vids2[:3])
    allv = vids2 or vids
    for pid in POSTS:
        try:
            req = urllib.request.Request(f"https://t.me/barook_ir/{pid}?embed=1&mode=tme", headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=40) as r:
                h = r.read().decode("utf-8", "replace")
            (RAW / f"proxy_embed_{pid}.html").write_text(h, encoding="utf-8")
            v = re.findall(r'<video[^>]+src="(https://[^"]+)"', h)
            if not v:
                v = re.findall(r'(https://cdn\d+\.telesco\.pe/file/[A-Za-z0-9_-]+)', h)
            print(f"post {pid} videos={len(v)} caption_has={('پرداخت' in h or 'آموزش' in h)}")
            if v:
                vurl = v[0]
                print(" downloading", vurl[:80])
                req = urllib.request.Request(vurl, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=120) as r:
                    blob = r.read()
                dest = VID / f"tg_{pid}.mp4"
                dest.write_bytes(blob)
                print(" saved", dest, len(blob))
        except Exception as e:
            print(f"post {pid} ERR", e)

print("working proxy:", working)
