# -*- coding: utf-8 -*-
"""Try downloading Telegram training videos via Playwright (system Chromium)."""
from pathlib import Path
import time, re, json

ROOT = Path(__file__).resolve().parent
VID = ROOT / "videos"
RAW = ROOT / "raw"
VID.mkdir(exist_ok=True)

POSTS = [108, 109, 110, 121, 122, 141, 19]

def main():
    from playwright.sync_api import sync_playwright
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            locale="fa-IR",
        )
        page = context.new_page()
        for pid in POSTS:
            url = f"https://t.me/barook_ir/{pid}?embed=1&mode=tme"
            info = {"id": pid, "url": url}
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=60000)
                time.sleep(2)
                html = page.content()
                (RAW / f"pw_embed_{pid}.html").write_text(html, encoding="utf-8")
                # find video src
                vids = re.findall(r'src="(https://cdn[^"]+\.mp4[^"]*)"', html)
                vids += re.findall(r'(https://cdn\d+\.telesco\.pe/file/[^"\']+)', html)
                info["video_candidates"] = vids[:5]
                caption = page.locator(".tgme_widget_message_text").inner_text() if page.locator(".tgme_widget_message_text").count() else ""
                info["caption"] = caption[:500]
                dur = page.locator(".tgme_widget_message_video_duration").inner_text() if page.locator(".tgme_widget_message_video_duration").count() else ""
                info["duration"] = dur
                # try download via page.request (same browser network stack)
                if vids:
                    vurl = vids[0]
                    try:
                        resp = page.request.get(vurl, timeout=120000)
                        data = resp.body()
                        dest = VID / f"tg_{pid}.mp4"
                        dest.write_bytes(data)
                        info["saved"] = str(dest)
                        info["size"] = len(data)
                        print(f"OK {pid} saved {len(data)}")
                    except Exception as e:
                        info["dl_err"] = str(e)
                        print(f"DL fail {pid}: {e}")
                else:
                    # try video element
                    vsrc = page.evaluate("""() => {
                      const v = document.querySelector('video');
                      return v ? (v.currentSrc || v.src || null) : null;
                    }""")
                    info["video_el"] = vsrc
                    if vsrc:
                        try:
                            resp = page.request.get(vsrc, timeout=120000)
                            data = resp.body()
                            dest = VID / f"tg_{pid}.mp4"
                            dest.write_bytes(data)
                            info["saved"] = str(dest)
                            info["size"] = len(data)
                            print(f"OK el {pid} saved {len(data)}")
                        except Exception as e:
                            info["dl_err"] = str(e)
                            print(f"DL el fail {pid}: {e}")
                    else:
                        print(f"NO VIDEO {pid} caption={caption[:80]!r}")
            except Exception as e:
                info["err"] = str(e)
                print(f"ERR {pid}: {e}")
            results.append(info)
        browser.close()
    (RAW / "pw_tg_results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print("done", len(results))

if __name__ == "__main__":
    main()
