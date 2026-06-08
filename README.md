# Cal Listener

Lean job-queue worker for [Cal Toolkit Web](https://cal-toolkit-web.vercel.app).
Runs on a Windows laptop with **Delivery Master** (and optionally
**ClearBooks** open in a browser). Pulls jobs from the Cal Toolkit Supabase
`job_queue` table, runs the matching handler, and writes results back.

## What it does

The web app is the UI. Anything that needs to drive the real Delivery
Master desktop window — or click around ClearBooks in a browser — gets
queued as a job here. The listener:

1. Registers itself in `listener_nodes` (heartbeats every 15s).
2. Polls `job_queue` for pending work.
3. Routes each job to the matching handler in `cal_listener/handlers/`.
4. Streams progress + log lines back to `job_progress`.
5. Marks the job `done` or `failed` when the handler returns.
6. Honours per-job cancel requests.

## Handlers implemented

| Handler key | What it does |
|---|---|
| `customer_360_sync`         | Walk DM customers, merge into Customer 360. |
| `customer_email_audit`      | Pull DM invoice/notes emails, diff against ClearBooks. |
| `dm_docket_search`          | Search DM for jobs by date / customer / status. |
| `revenue_breakdown_scraper` | Scrape DM revenue breakdown screens. |
| `tariff_retrigger_dry_run`  | Walk BT refs, record would-be tariff changes. |
| `dm_daily_check`            | Pull yesterday's DM jobs into the daily-check table. |
| `invoice_plan_run`          | Drive DM through customer + PO to generate invoices. |
| `cb_create_bill`            | Raise a supplier bill in ClearBooks. |
| `cb_edit_bill`              | Edit an existing ClearBooks bill. |
| `cb_credit_note`            | Issue a credit note. |
| `cb_mark_bill_paid`         | Mark one or more bills as paid. |
| `cb_money_in_out`           | Record money received or spent. |
| `cb_statements`             | Generate customer statements. |

## Install

**Download `CalListenerSetup.exe`** from the [Releases](../../releases)
page and double-click. The installer:

- Bundles Python, pywinauto and all dependencies — nothing else to install.
- Asks you for the listener id (e.g. `listener-a`), the Supabase service
  key, and your DM credentials.
- Registers itself as a Scheduled Task that runs at logon and auto-restarts
  on crash.
- That's it.

No zip files. No PowerShell. No `pip install`.

## Dev install (running from source)

```cmd
git clone https://github.com/OwenJohnson505/cal-listener
cd cal-listener
python -m venv venv
venv\Scripts\activate
pip install -e .
copy cal_listener\secrets.example.json cal_listener\secrets.json
:: edit secrets.json
python -m cal_listener
```

## Building the .exe locally

```cmd
pip install pyinstaller
python installer\build_exe.py
:: outputs dist\CalListenerSetup.exe
```

## Architecture

```
cal_listener/
├── daemon.py            # main loop: register, heartbeat, claim, dispatch
├── supabase.py          # minimal HTTP wrapper round PostgREST + RPC
├── secrets.py           # secrets loader + first-run config dialog
├── dm.py                # Delivery Master driver (pywinauto)
├── cb.py                # ClearBooks driver (browser via Playwright)
├── handlers/            # one file per job type — each exposes run()
│   ├── __init__.py      # HANDLERS registry
│   ├── customer_360_sync.py
│   ├── ...
└── secrets.example.json
installer/
├── build_exe.py         # PyInstaller wrapper
├── first_run_dialog.py  # tkinter GUI shown on first launch
└── installer.iss        # optional Inno Setup wrapper for an MSI feel
.github/workflows/
└── build.yml            # auto-build the .exe on push to main
```

## Cancellation, progress, errors

Every handler receives an `on_progress(message, percent=None, level='info',
detail=None)` callable. Inside the handler, call it whenever you want a
log line + percent bar update to flow to the web. If the user clicked
"Cancel" in the web, the next `on_progress` call raises `JobCancelled`,
which the daemon catches and marks the job cancelled. No polling required
inside the handler.

Exceptions bubble out of `run()` — the daemon catches them, captures the
traceback into `job_queue.error`, marks the job failed, and keeps running.
