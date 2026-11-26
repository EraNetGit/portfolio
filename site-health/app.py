"""Site Health Dashboard (demo build).

Reads a monitoring "master" workbook (one tab per report type) and turns it
into a queue of sites that need attention, ranked by how often and how
severely they have been flagged recently. Ships with synthetic sample data.

    python app.py   ->  http://127.0.0.1:5002
"""
import os
from collections import Counter, defaultdict
from datetime import date, timedelta
from pathlib import Path

from flask import Flask, jsonify, render_template
from openpyxl import load_workbook

DATA = Path(os.environ.get("DEMO_XLSX", Path(__file__).parent / "data" / "demo_monitoring_master.xlsx"))
WINDOW_DAYS = 14
VISIT_DAYS = 4      # flagged on this many distinct days -> recommend a visit
WATCH_DAYS = 2      # ...or this many -> keep watching

app = Flask(__name__)


def load_events():
    wb = load_workbook(DATA, read_only=True, data_only=True)
    sites, events = {}, []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        head = [str(h) for h in rows[0]]
        for r in rows[1:]:
            rec = dict(zip(head, r))
            if ws.title == "Sites":
                sites[rec["code"]] = rec
            elif rec.get("date"):
                rec["report"] = ws.title
                events.append(rec)
    return sites, events


def build_queue():
    sites, events = load_events()
    as_of = max(date.fromisoformat(str(e["date"])[:10]) for e in events)
    days = [as_of - timedelta(days=i) for i in range(WINDOW_DAYS - 1, -1, -1)]
    start = days[0]

    per_site = defaultdict(list)
    for e in events:
        d = date.fromisoformat(str(e["date"])[:10])
        if d >= start:
            per_site[e["client_code"]].append((d, e))

    cards = []
    for code, items in per_site.items():
        distinct = {d for d, _ in items}
        score = sum(int(e["severity"]) for _, e in items)
        by_event = Counter(e["event"] for _, e in items)
        top_event, top_n = by_event.most_common(1)[0]
        if len(distinct) >= VISIT_DAYS:
            action, level = "Recommend tech visit", "visit"
        elif len(distinct) >= WATCH_DAYS:
            action, level = "Keep watching", "watch"
        else:
            action, level = "No action", "ok"
        per_day = Counter(d for d, _ in items)
        site = sites.get(code, {})
        cards.append({
            "code": code,
            "client": items[0][1]["client"],
            "provider": site.get("monitoring") or "Unmonitored",
            "address": site.get("address", ""),
            "score": score,
            "days_flagged": len(distinct),
            "events": len(items),
            "top_event": top_event,
            "top_event_n": top_n,
            "breakdown": [{"event": k, "n": v} for k, v in by_event.most_common()],
            "reports": sorted({e["report"] for _, e in items}),
            "spark": [per_day.get(d, 0) for d in days],
            "action": action,
            "level": level,
        })
    cards.sort(key=lambda c: (c["level"] != "visit", c["level"] != "watch", -c["score"]))
    return {
        "as_of": as_of.strftime("%d %b %Y"),
        "window": WINDOW_DAYS,
        "day_labels": [d.strftime("%d %b") for d in days],
        "counts": Counter(c["level"] for c in cards),
        "cards": [c for c in cards if c["level"] != "ok"],
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/queue")
def api_queue():
    return jsonify(build_queue())


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5002, debug=False)
