# -*- coding: utf-8 -*-
"""Import publicly listed Telegram group/channel admin mobiles from hop graph."""
from __future__ import annotations

import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "batches" / "person_id_hits.jsonl"
GRAPH = ROOT / "raw" / "telegram_hop_graph.json"

HITS = [
    {
        "company": "تالار پخش مواد غذایی و بهداشتی",
        "person": "حقیقی — مدیرگروه تلگرام",
        "mobile": "09124267508",
        "source": "telegram_hop:talarpakhsh_bio",
    },
    {
        "company": "تالار پخش مواد غذایی و بهداشتی",
        "person": "واهب — مدیرگروه تلگرام",
        "mobile": "09131030997",
        "source": "telegram_hop:talarpakhsh_bio",
    },
    {
        "company": "تالار پخش مواد غذایی و بهداشتی",
        "person": "جعفری — مدیرگروه تلگرام",
        "mobile": "09361579367",
        "source": "telegram_hop:talarpakhsh_bio",
    },
    {
        "company": "انجمن صنفی روابط عمومی تهران",
        "person": "روابط عمومی انجمن",
        "mobile": "09011016447",
        "source": "telegram_hop:@pra_ir",
    },
    {
        "company": "مدیران کارآفرین",
        "person": "وحید حامد — مدیر کانال",
        "mobile": "09153062245",
        "source": "telegram_hop:@modiran_karafarin",
    },
    {
        "company": "انجمن صنفی مدیران و مشاورین منابع انسانی",
        "person": "ادمین انجمن",
        "mobile": "09123843135",
        "source": "telegram_hop:@hrmSupergroup",
    },
]


def main() -> None:
    existing = []
    if OUT.exists():
        existing = [json.loads(l) for l in OUT.read_text(encoding="utf-8").splitlines() if l.strip()]
    seen_mobs = {h.get("mobile") for h in existing if h.get("mobile")}
    added = 0
    with OUT.open("a", encoding="utf-8") as f:
        for h in HITS:
            mob = h["mobile"]
            if mob in seen_mobs:
                continue
            f.write(json.dumps(h, ensure_ascii=False) + "\n")
            seen_mobs.add(mob)
            added += 1
    print(json.dumps({"added": added, "graph": str(GRAPH.exists())}, ensure_ascii=False))


if __name__ == "__main__":
    main()
