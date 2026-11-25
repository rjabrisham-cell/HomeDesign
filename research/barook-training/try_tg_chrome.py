# -*- coding: utf-8 -*-
from pathlib import Path
import time, re, json

ROOT = Path(__file__).resolve().parent
RAW, VID = ROOT / "raw", ROOT / "videos"
POSTS = [108, 109, 110, 121, 122, 141]

def main():
    from playwright.sync_api import sync_playwright
    results = []
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            channel="chrome",
            args=["--disable-blink-features=AutomationControlled"],
        )
        page = browser.new_page()
        for pid in POSTS:
            url = f"https://t.me/barook_ir/{pid}?embed=1&mode=tme"
            info = {"id": pid}
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=90000)
                time.sleep(3)
                html = page.content()
                (RAW / f"chrome_embed_{pid}.html").write_text(html, encoding="utf-8")
                vsrc = page.evaluate("""() => {
                  const v=document.querySelector('video');
                  return v ? (v.currentSrc||v.src||null) : null;
                }""")
                info["vsrc"] = vsrc
                cap = ""
                if page.locator(".tgme_widget_message_text").count():
                    cap = page.locator(".tgme_widget_message_text").first.inner_text()
                info["caption"] = cap[:400]
                print(f"{pid} vsrc={bool(vsrc)} cap={cap[:60]!r}")
                if vsrc:
                    resp = page.request.get(vsrc, timeout=180000)
                    blob = resp.body()
                    dest = VID / f"tg_{pid}.mp4"
                    dest.write_bytes(blob)
                    info["size"] = len(blob)
                    print(" saved", len(blob))
            except Exception as e:
                info["err"] = str(e)
                print("ERR", pid, e)
            results.append(info)
        browser.close()
    (RAW / "chrome_tg_results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")

if __name__ == "__main__":
    main()
