# DipSeva Rental — Ledger + Dashboard (Marathi)

New clean rental accounting for Mandap / Caterers. Tally fails on multi-day entries — this solves it.

## What is this?
- One entry book with 3 types: `Issue` (सामान दिले) / `Return` (परत आले) / `Payment` (पैसे जमा)
- Running khata per customer across many days
- Example Yogesh: Issue 850 + Paid 600 → pending 250. Next Issue 900 + Paid 0 → total pending **1150**. Next Payment 500 → 650.
- Old `Ledger JULY 2026.xlsx` bugs fixed: double-minus balance (550→1150), inverted return logic, stock ignoring opening, dashboard pending sign, 4 hard-coded totals.

## Repo layout
```
Dipseva_Ledger_NEW.xlsx      # real database (Sheet-first, works alone)
dipseva_app/
  app.py                     # Flask dashboard — ALL entry from here
  store.py / summary.py      # read sheet, compute totals (single logic)
  templates/                 # Mukta + Noto Sans Devanagari UI
  static/app.css / app.js
  requirements.txt run.sh README.md
tools (build_workbook split):
  dipseva_app/b0.py..b9.py   # workbook builder (Ledger/Items/Customers/Dashboard/Khata/Stock/Statement/Help + migration)
```

## Run (no 24/7 PC needed for data — sheet is DB)
```
cd dipseva_app
pip install -r requirements.txt
python3 app.py
# open http://127.0.0.1:5000
# phone same WiFi: http://<shop-PC-IP>:5000
```

## Pages
- `/` Home — business-date presets, 4 KPIs, top receivables (click row → statement, ₹ icon → record payment), 10-day collections, recent activity, low stock
- `/entry` — type cards (Issue / Return / Payment) that reshape the form, quick item chips, qty stepper, live summary, optional "keep form open"
- `/customers` — KPIs + instant search + sortable columns + pending-% meter + row actions
- `/customer/<name>` — statement with running balance, collect-payment panel, WhatsApp reminder, per-customer phone save, print/PDF
- `/stock` — low-stock KPIs, health meters, status badges, filter + only-low toggle
- `/api/summary` — JSON

## UI / UX
- Warm Ivory theme: ivory `#FAF7F0` bg, white cards, deep navy `#1E2A5A`, maroon accent, serif headings — no gradients/neon
- Toasts for save/error feedback, clickable rows, sortable + sticky table headers, live filters (`/` to focus, `Esc` to clear)
- Keyboard: `Alt+1` overview, `Alt+2` entry, `Alt+3` customers, `Alt+4` inventory
- Prefill deep links: `/entry?customer=X&type=Payment&amount=Y`, `/entry?item=CH&type=Issue`
- Print stylesheet on the statement (sidebar/actions hidden on paper)
- Fonts: Inter + Mukta (Marathi names/notes)
- Sheet tab colours match web cards

## Note
- Entry only from the app. Sheet is backup if issue happens.
- Old file `Ledger JULY 2026.xlsx` is reference only, not used by app.
- This repo includes `*.xlsx` with customer data. Make PRIVATE if you add real phones.
- openpyxl gotcha: `ws.cell(r, c, None)` does **not** blank a cell — it ignores `None`. Use `ws.cell(r, c).value = None`.
