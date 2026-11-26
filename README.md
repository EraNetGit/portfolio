# Operations tooling portfolio

Small internal-tool demos for service and operations teams. Everything here runs on
**synthetic sample data** for a fictional company ("Kestrel Bay Systems").

| Project | What it does | Run |
|---|---|---|
| [`task-dashboard`](task-dashboard) | Task pipeline by status, technician workload, status-vs-age heatmap, stale-task list | `python app.py` -> http://127.0.0.1:5001 |
| [`site-health`](site-health) | Ranked queue of sites that keep faulting, with a recommended next action | `python app.py` -> http://127.0.0.1:5002 |

Windows: double-click `run.bat` inside either folder (creates a venv on first run).

Using your own data: set `DEMO_XLSX` to a workbook with the same columns as the sample in `data/`.

## Layout

- `demo_data/generate.py`: seeded generator that writes the sample workbooks (`python demo_data/generate.py`)
- `site/`: the static portfolio page, published with GitHub Pages
- `scripts/check_no_real_data.py`: fails CI if forbidden names or secrets appear (edit `FORBIDDEN` first)
