# -*- coding: utf-8 -*-
from pathlib import Path
import re, json
html = Path(r"D:\HomeDesign\research\barook-kasbno\raw\barook_fmcg.html").read_text(encoding="utf-8")
texts = re.findall(r"[\u0600-\u06FF][^\n<\"]{8,160}", html)
seen = set()
out = []
keys = ["ثبت", "پرداخت", "پوز", "کارت", "فاکتور", "شعبه", "چک", "بازپرداخت", "مرحله", "مراجعه", "کشید", "اعتبار", "سقف", "کارمزد"]
for t in texts:
    t = t.strip()
    if t in seen:
        continue
    if any(k in t for k in keys):
        seen.add(t)
        out.append(t)
Path(r"D:\HomeDesign\research\barook-training\raw\fmcg_steps.json").write_text(
    json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8"
)
print("n=", len(out))
for t in out[:100]:
    print("-", t)
