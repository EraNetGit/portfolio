"""Workflow Task Dashboard (demo build).

Reads a tasks spreadsheet and shows the pipeline by status, technician
workload, and how stale each open status is. Ships with synthetic sample
data; point DEMO_XLSX at another file with the same columns to use your own.

    python app.py   ->  http://127.0.0.1:5001
"""
import os
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, render_template
from openpyxl import load_workbook

DATA = Path(os.environ.get("DEMO_XLSX", Path(__file__).parent / "data" / "demo_tasks.xlsx"))
STATUS_ORDER = ["New", "Quoted", "Scheduled", "In Progress", "Awaiting Parts", "Awaiting Client", "Completed"]
BUCKETS = [("0-1d", 0, 1), ("2-3d", 2, 3), ("4-7d", 4, 7), ("8-14d", 8, 14), ("15d+", 15, 10**6)]
STALE_DAYS = 7

app = Flask(__name__)


def load_tasks():
    ws = load_workbook(DATA, read_only=True, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    head = [str(h) for h in rows[0]]
    return [dict(zip(head, r)) for r in rows[1:] if any(r)]


def summarise():
    tasks = load_tasks()
    for t in tasks:
        t["updated_dt"] = datetime.fromisoformat(str(t["updated"]))
    as_of = max(t["updated_dt"] for t in tasks)  # keeps the demo stable over time
    for t in tasks:
        t["age"] = (as_of - t["updated_dt"]).days
    open_tasks = [t for t in tasks if t["status"] != "Completed"]

    pipeline = Counter(t["status"] for t in tasks)
    workload = Counter(t["tech"] or "Unassigned" for t in open_tasks)

    open_statuses = [s for s in STATUS_ORDER if s != "Completed"]
    grid = {s: [0] * len(BUCKETS) for s in open_statuses}
    for t in open_tasks:
        for i, (_, lo, hi) in enumerate(BUCKETS):
            if lo <= t["age"] <= hi:
                grid[t["status"]][i] += 1

    stale = sorted((t for t in open_tasks if t["age"] > STALE_DAYS), key=lambda t: -t["age"])
    return {
        "as_of": as_of.strftime("%d %b %Y"),
        "kpi": {
            "open": len(open_tasks),
            "stale": len(stale),
            "awaiting": sum(1 for t in open_tasks if t["status"].startswith("Awaiting")),
            "completed": pipeline.get("Completed", 0),
        },
        "pipeline": [{"status": s, "count": pipeline.get(s, 0)} for s in STATUS_ORDER],
        "workload": [{"tech": k, "count": v} for k, v in sorted(workload.items(), key=lambda kv: -kv[1])],
        "heatmap": {"buckets": [b[0] for b in BUCKETS], "rows": [{"status": s, "cells": grid[s]} for s in open_statuses]},
        "stale": [
            {"id": t["id"], "client": t["client"], "title": t["title"], "status": t["status"],
             "tech": t["tech"] or "Unassigned", "age": t["age"]}
            for t in stale[:15]
        ],
        "stale_days": STALE_DAYS,
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/summary")
def api_summary():
    return jsonify(summarise())


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5001, debug=False)
