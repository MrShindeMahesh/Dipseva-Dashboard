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
- `/` Home — Billed / Received / Pending / Today / Top due / 10-day collection / Recent / Low stock
- `/entry` — new Issue / Return / Payment (rate auto, bill live)
- `/customers`, `/customer/<name>` — khata + WhatsApp copy text
- `/stock` — InHand = Opening + Return − Given
- `/api/summary` — JSON

## Design
- Fonts: Mukta + Noto Sans Devanagari (Marathi must)
- Colours: Navy #1B2A4E, Saffron #F59E0B/#D97706, Teal #0D9488
- Sheet tab colours match web cards

## Note
- Entry only from Dashboard. Sheet is backup if issue happens.
- Old file `Ledger JULY 2026.xlsx` is reference only, not used by app.
- This repo includes `*.xlsx` with customer data. Make PRIVATE if you add real phones.
