"""Synthetic demo data for the portfolio dashboards.

Everything here is invented: the company ("Kestrel Bay Systems"), clients,
technicians, addresses and events. No real data is read or derived.
Seeded so every run produces the same dataset.

Usage: python generate.py  -> writes tasks.json, sites.json, events.json
"""
import json
import random
from datetime import datetime, timedelta
from pathlib import Path

SEED = 20261126
random.seed(SEED)
NOW = datetime(2026, 11, 26, 9, 0)
OUT = Path(__file__).parent

TECHS = ["Sam Okafor", "Priya Lindqvist"]

CLIENT_NAMES = [
    "Ashgrove Dental", "Bluegum Cafe", "Corvus Logistics", "Delmar Apartments",
    "Eastfield Veterinary", "Fernbank Primary", "Gantry Print Co", "Harlow Physio",
    "Ironbark Storage", "Juniper Hair Studio", "Kelvin Auto Group", "Lanternfish Bistro",
    "Marlow Accounting", "Nettle & Co Florist", "Oakridge Church Hall", "Pelican Marine Supplies",
    "Quarry Lane Gym", "Redwater Pharmacy", "Saltbush Childcare", "Tallow Bakery",
    "Umber Architects", "Vesper Wine Bar", "Wattle Street Medical", "Yarrow Retirement Village",
    "Zephyr Cycles", "Birchwood Law", "Cobalt Data Rooms", "Driftwood Motel",
    "Emberly Bookshop", "Foxglove Nursery",
]
STREETS = ["Example St", "Sample Rd", "Placeholder Ave", "Demo Pde", "Test Crt", "Mock Lane"]
SUBURBS = ["Northfield", "Eastvale", "Port Alder", "Westmoor", "Southbrook", "Midhaven"]

JOB_TYPES = [
    ("CCTV camera offline", "CCTV"),
    ("NVR storage check", "CCTV"),
    ("Alarm keypad fault", "Alarm"),
    ("Panel battery replacement", "Alarm"),
    ("Access reader not working", "Access"),
    ("WiFi dropouts in back office", "Network"),
    ("Switch replacement", "Network"),
    ("Annual service visit", "Service"),
    ("New camera install quote", "Quote"),
    ("Intercom audio issue", "AV"),
]
STATUSES = ["New", "Quoted", "Scheduled", "In Progress", "Awaiting Parts", "Awaiting Client", "Completed"]
STATUS_W = [10, 8, 18, 14, 8, 10, 32]


def client_code(i):
    return f"DEMO{i + 1:03d}"


def make_sites():
    sites = []
    for i, name in enumerate(CLIENT_NAMES):
        sites.append({
            "code": client_code(i),
            "name": name,
            "address": f"{random.randint(1, 199)} {random.choice(STREETS)}, {random.choice(SUBURBS)}",
            "monitoring": random.choice(["Provider A", "Provider B", "Provider C", None]),
        })
    return sites


def make_tasks(sites):
    tasks = []
    for n in range(1, 91):
        site = random.choice(sites)
        title, category = random.choice(JOB_TYPES)
        status = random.choices(STATUSES, STATUS_W)[0]
        created = NOW - timedelta(days=random.randint(0, 60), hours=random.randint(0, 8))
        age_since_update = random.choice([0, 0, 1, 1, 2, 3, 5, 8, 12, 20])
        updated = min(NOW, created + timedelta(days=random.randint(0, 30)))
        if status != "Completed":
            updated = NOW - timedelta(days=age_since_update, hours=random.randint(0, 6))
            updated = max(updated, created)
        tasks.append({
            "id": f"T{1000 + n}",
            "title": title,
            "category": category,
            "client_code": site["code"],
            "client": site["name"],
            "status": status,
            "tech": random.choice(TECHS + [None]) if status in ("New", "Quoted") else random.choice(TECHS),
            "created": created.isoformat(timespec="minutes"),
            "updated": updated.isoformat(timespec="minutes"),
        })
    return tasks


EVENT_TYPES = [
    ("No Signals Received from Client", "Delinquency", 3),
    ("Alarm Repeated Loss of Polling", "Delinquency", 3),
    ("Battery Test Failure", "System Trouble", 2),
    ("Sensor Battery Low", "Alarm Events", 2),
    ("Zone Isolate", "Alarm Events", 1),
    ("Panel Serial Connection Issue", "System Trouble", 2),
    ("TICKET FOR DEALER", "Incident", 1),
]


def make_events(sites):
    """30 days of events. A handful of sites get recurring issues so the
    'needs attention' queue has realistic, repeating patterns."""
    problem_sites = random.sample(sites, 7)
    events = []
    for day in range(30):
        date = (NOW - timedelta(days=day)).date()
        for s in problem_sites:
            if random.random() < 0.55:
                et, report, sev = random.choice(EVENT_TYPES)
                events.append({"date": date.isoformat(), "client_code": s["code"],
                               "client": s["name"], "event": et, "report": report, "severity": sev})
        # background noise on other sites
        for s in random.sample(sites, 3):
            et, report, sev = random.choice(EVENT_TYPES)
            events.append({"date": date.isoformat(), "client_code": s["code"],
                           "client": s["name"], "event": et, "report": report, "severity": sev})
    return events


if __name__ == "__main__":
    sites = make_sites()
    tasks = make_tasks(sites)
    events = make_events(sites)
    for name, data in (("sites", sites), ("tasks", tasks), ("events", events)):
        (OUT / f"{name}.json").write_text(json.dumps(data, indent=1))
    print(f"sites={len(sites)} tasks={len(tasks)} events={len(events)}")


def write_xlsx():
    """Write the same synthetic data as Excel workbooks for the local demos."""
    from collections import defaultdict
    from openpyxl import Workbook

    sites = json.loads((OUT / "sites.json").read_text())
    tasks = json.loads((OUT / "tasks.json").read_text())
    events = json.loads((OUT / "events.json").read_text())

    def sheet(ws, rows, headers):
        ws.append(headers)
        for r in rows:
            ws.append([r.get(h) for h in headers])
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(
                40, max(len(str(c.value or "")) for c in col) + 2)
        ws.freeze_panes = "A2"

    wb = Workbook()
    ws = wb.active
    ws.title = "Tasks"
    sheet(ws, tasks, ["id", "title", "category", "client_code", "client", "status", "tech", "created", "updated"])
    wb.save(OUT / "demo_tasks.xlsx")

    wb = Workbook()
    wb.remove(wb.active)
    by_report = defaultdict(list)
    for e in events:
        by_report[e["report"]].append(e)
    for report, rows in by_report.items():
        sheet(wb.create_sheet(report[:31]), rows, ["date", "client_code", "client", "event", "severity"])
    ws = wb.create_sheet("Sites")
    sheet(ws, sites, ["code", "name", "address", "monitoring"])
    wb.save(OUT / "demo_monitoring_master.xlsx")
    print("wrote demo_tasks.xlsx, demo_monitoring_master.xlsx")


write_xlsx()
