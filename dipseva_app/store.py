"""Workbook IO + summaries for Dipseva."""
import os
from datetime import datetime, date, timedelta
import openpyxl
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
XLSX = os.path.join(BASE, "Dipseva_Ledger_NEW.xlsx")
def load_wb():
    return openpyxl.load_workbook(XLSX)
def parse_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%d %b %Y"):
        try:
            return datetime.strptime(str(v).strip(), fmt).date()
        except Exception:
            pass
    try:
        return datetime.fromisoformat(str(v)).date()
    except Exception:
        return None
def read_ledger():
    wb = load_wb()
    ws = wb["Ledger"]
    wi = wb["Items"]
    imap = {}
    for rr in range(3, 203):
        k = wi.cell(rr, 1).value
        if k not in (None, ""):
            imap[str(k).strip()] = str(wi.cell(rr, 2).value or "").strip()
    rows = []
    for r in range(3, 1003):
        a = ws.cell(r, 1).value; b = ws.cell(r, 2).value; c = ws.cell(r, 3).value
        if (a in (None, "")) and (b in (None, "")) and (c in (None, "")):
            continue
        d = ws.cell(r, 4).value; f = ws.cell(r, 6).value; g = ws.cell(r, 7).value
        i = ws.cell(r, 9).value
        try: qty = float(f or 0)
        except Exception: qty = 0
        try: rate = float(g or 0)
        except Exception: rate = 0
        typ = str(c or "").strip()
        bill = qty * rate if typ == "Issue" else 0.0
        try: paid = float(i or 0)
        except Exception: paid = 0.0
        dt = parse_date(a)
        item = str(d or "").strip()
        if item.startswith("="):
            item = ""
        ev = ws.cell(r, 5).value
        es = str(ev or "").strip()
        if es.startswith("=") or es in ("", "ERR"):
            iname = imap.get(item, "")
        else:
            iname = es
        rows.append({"row": r, "date": dt.isoformat() if dt else "",
            "date_disp": dt.strftime("%d-%b") if dt else "",
            "customer": str(b or "").strip(), "type": typ,
            "item": item,
            "item_name": iname,
            "qty": int(qty) if qty == int(qty) else qty,
            "rate": rate, "bill": bill, "paid": paid,
            "mode": str(ws.cell(r, 10).value or "").strip(),
            "note": str(ws.cell(r, 11).value or "").strip()})
    return rows
def read_masters():
    wb = load_wb()
    wi = wb["Items"]; wc = wb["Customers"]
    items = []
    for r in range(3, 203):
        a = wi.cell(r, 1).value
        if a in (None, ""):
            continue
        items.append({"id": str(a).strip(), "name": str(wi.cell(r, 2).value or "").strip(),
            "marathi": str(wi.cell(r, 3).value or "").strip(),
            "rate": wi.cell(r, 4).value or 0, "opening": wi.cell(r, 5).value or 0,
            "min": wi.cell(r, 6).value or 0})
    custs = []
    for r in range(3, 203):
        a = wc.cell(r, 1).value
        if a in (None, ""):
            continue
        custs.append({"name": str(a).strip(), "person": str(wc.cell(r, 2).value or "").strip(),
            "phone": str(wc.cell(r, 3).value or "").strip(),
            "ctype": str(wc.cell(r, 4).value or "").strip(),
            "addr": str(wc.cell(r, 5).value or "").strip(),
            "opening_due": wc.cell(r, 6).value or 0})
    return items, custs
