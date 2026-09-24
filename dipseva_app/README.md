# DipSeva Rental — Ledger + Dashboard

New clean design. Old file `Ledger JULY 2026.xlsx` is NOT touched.

## Files
- `Dipseva_Ledger_NEW.xlsx` — the real database (Sheet-first).
  Sheets: `Ledger | Items | Customers | Dashboard | Khata | Stock | Statement | Help`
- `dipseva_app/app.py` — web dashboard (all entry from here).
- `dipseva_app/store.py + summary.py` — read sheet, compute totals.

## Logic (multi-day, your Tally problem solved)
One book, three types:
- `Issue` = सामान दिले → Bill = Qty × Rate (auto), Paid = त्यादिवशी जमा
- `Return` = परत आले → फक्त Qty (Bill 0)
- `Payment` = नंतर पैसे → फक्त Paid

Yogesh example:
- 1 Jul Issue 850 + Paid 600 → pending 250
- 5 Jul Issue 900 + Paid 0 → total pending **1150** (correct, single minus)
- Later Payment 500 → pending 650

Khata pending = Billed + OpeningDue − Received (ONE minus, old bug of double minus fixed).
Stock InHand = Opening + Returned − Given (old E−F fixed).
Dashboard Pending = Billed − Received (old Received−Billed fixed).

## Run
```
cd dipseva_app
python3 app.py
# open http://127.0.0.1:5000
```
From phone on same WiFi: `http://<shop-PC-IP>:5000`

## Rules
- Entry only from Dashboard. Sheet is backup for problems.
- Marathi font: Mukta + Noto Sans Devanagari (web) / Noto Sans Devanagari (sheet).
- Colours: Navy #1B2A4E, Saffron #F59E0B/#D97706, Teal #0D9488, Green/Red for status.

## Fixes migrated from old sheet
- 4 hard TOTALs recomputed (TBL 1000→780, DSH 540→400, DRM 120→20, DSH 250→200), note kept in Ledger Note.
- Old RETURN QTY < QTY rows split into Return rows.
- Advance row (Atul Salve 1000, no item) → Payment row.
