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
- `/` Home — business-date presets, 4 KPIs, top receivables (click row → statement, ₹ icon → record payment), 10-day collections (row → that day in the calendar), recent activity with ✎ edit, low stock
- `/entry` — pick customer + date once, then add many **item lines** in one save; each line has its own **Give ↑ / Take ↓** toggle, qty stepper, rate *and* a typeable **Amount ₹** (amount wins; rate is recalculated). Paid / Mode / Note are shared. Buttons: *Save*, *Save & open statement*, *Keep form open*, plus a “last entries of this customer” list with ✎ edit
- `/customers` — KPIs + instant search + sortable columns + pending-% meter + row actions
- `/customer/<name>` — statement with running balance, collect-payment panel, WhatsApp reminder, per-customer phone save, print/PDF, ✎ edit on every ledger row
- `/stock` — low-stock KPIs, health meters, status badges, filter + only-low toggle
- `/calendar` — month grid (bill ₹, customer count, ↑ given / ↓ returned per day); tapping a date lists **each customer of that day with the stock lines below** (qty, rate, bill, paid, note, ✎ edit) + “+ Entry on this date”
- `/entry/<row>/edit` — change a saved line: date, customer, type, item, qty, rate, amount, paid, mode, note (+ safe delete). Row number = Excel row
- `/api/summary` — JSON

## UI / UX
- Warm Ivory theme: ivory `#FAF7F0` bg, white cards, deep navy `#1E2A5A`, maroon accent, serif headings — no gradients/neon
- Money is always **₹** (KPI, tables, share text, and the rupee icon on “record payment” buttons). Amount boxes are free text: `1150`, `1,150`, `₹1,150`, `Rs 1150` all parse
- Toasts for save/error feedback, clickable rows, sortable + sticky table headers, live filters (`/` to focus, `Esc` to clear)
- Keyboard: `Alt+1` overview, `Alt+2` entry, `Alt+3` customers, `Alt+4` inventory, `Alt+5` calendar
- Prefill deep links: `/entry?customer=X&type=Payment&amount=Y`, `/entry?item=CH&type=Issue`, `/entry?date=2026-07-05`
- Layout fills the screen (no 1180px cap) — wide monitors use the full width
- Mobile (performance-first, CSS only — no extra JS libraries): tables become labelled cards, 44px touch targets, 16px inputs (no iOS zoom), sticky save bar, compact calendar, toasts at the bottom
- Print stylesheet on the statement (sidebar/actions hidden on paper)
- Fonts: Inter + Mukta (Marathi names/notes)
- Sheet tab colours match web cards

## Note
- Entry only from the app. Sheet is backup if issue happens.
- Multiple item lines in one save = one Excel row per line (same date/customer/note); Paid/Mode land on the first line only.
- Changing a row’s date later is expected: every ledger row is editable from the statement, dashboard, calendar or the entry page (`✎`).
- Old file `Ledger JULY 2026.xlsx` is reference only, not used by app.
- This repo includes `*.xlsx` with customer data. Make PRIVATE if you add real phones.
- openpyxl gotcha: `ws.cell(r, c, None)` does **not** blank a cell — it ignores `None`. Use `ws.cell(r, c).value = None`.
- `$` inside Ledger formulas (`$D3`, `$F3`) is Excel’s **absolute reference**, not currency — leave it alone.
