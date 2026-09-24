import calendar as pycal
from datetime import date, timedelta
from urllib.parse import quote
from flask import Flask, render_template, request, redirect, url_for, jsonify, flash
from store import (XLSX, load_wb, parse_date, read_ledger, read_masters, money,
    next_free_row, write_line, clear_row, MODES, TYPES)
from summary import summarize, day_map, day_groups, month_stats
app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = "dipseva-local-2026"


def lines_from_form(f, default_type):
    """Multi-item rows: l_type[] / l_item[] / l_qty[] / l_rate[] / l_amount[].
    Amount wins over rate so a typed lump-sum is what reaches the sheet."""
    ts = f.getlist("l_type"); its = f.getlist("l_item")
    qs = f.getlist("l_qty"); rs = f.getlist("l_rate"); ams = f.getlist("l_amount")
    n = max(len(ts), len(its), len(qs), len(rs), len(ams))
    out = []
    for i in range(n):
        typ = (ts[i] if i < len(ts) else "").strip() or default_type
        if typ not in ("Issue", "Return"):
            typ = "Issue"
        item = (its[i] if i < len(its) else "").strip()
        qty = money(qs[i] if i < len(qs) else "", 0)
        rate = money(rs[i] if i < len(rs) else "", 0)
        amt = money(ams[i] if i < len(ams) else "", 0)
        if amt and qty:
            rate = round(amt / qty, 2)
        if not item:
            continue
        if not qty and not rate:
            continue
        out.append({"type": typ, "item": item, "qty": qty, "rate": rate, "amount": qty * rate})
    return out


@app.route("/")
def home():
    today = request.args.get("today") or date.today().isoformat()
    try: base_day = date.fromisoformat(today)
    except Exception: base_day = date.today()
    presets = [("Today", base_day.isoformat()),
        ("Yesterday", (base_day - timedelta(days=1)).isoformat()),
        ("3 days back", (base_day - timedelta(days=3)).isoformat()),
        ("A week back", (base_day - timedelta(days=7)).isoformat())]
    rows = read_ledger(); items, custs = read_masters()
    s = summarize(rows, items, custs, today)
    recent = sorted([r for r in rows if r["date"]], key=lambda x: x["date"], reverse=True)[:15]
    return render_template("dashboard.html", s=s, items=items, custs=custs,
        recent=recent, today=today, presets=presets)
@app.route("/entry", methods=["GET", "POST"])
def entry():
    items, custs = read_masters()
    if request.method == "POST":
        f = request.form
        dt = parse_date(f.get("date")) or date.today()
        cust = f.get("customer", "").strip()
        typ = (f.get("type", "Issue") or "Issue").strip()
        if typ not in TYPES:
            typ = "Issue"
        mode = f.get("mode", "").strip()
        note = f.get("note", "").strip()
        paid = money(f.get("paid"), 0)
        back = url_for("entry", customer=cust, type=typ)
        if not cust:
            flash("Customer निवडा — नाव लिहा किंवा यादीतून घ्या", "err")
            return redirect(url_for("entry"))
        lines = lines_from_form(f, typ)
        if typ == "Payment":
            if paid <= 0:
                flash("रक्कम लिहा — जमा झालेली amount टाका", "err")
                return redirect(back)
        elif not lines:
            flash("Item निवडा — किमान एक वस्तू द्या किंवा घ्या", "err")
            return redirect(back)
        wb = load_wb(); ws = wb["Ledger"]
        rmap = {i["id"]: money(i["rate"], 0) for i in items}
        written, bill, given, got = 0, 0.0, 0.0, 0.0
        if typ == "Payment":
            r = next_free_row(ws)
            if r is None:
                flash("Ledger full — नवीन sheet लागेल", "err")
                return redirect(back)
            write_line(ws, r, dt, cust, "Payment", paid=paid, mode=mode, note=note)
            written = 1
        else:
            for i, ln in enumerate(lines):
                r = next_free_row(ws)
                if r is None:
                    flash("Ledger full — %d line जतन झाल्या" % written, "err")
                    break
                ln_rate = ln["rate"] or rmap.get(ln["item"], 0)
                write_line(ws, r, dt, cust, ln["type"], item=ln["item"], qty=ln["qty"],
                    rate=ln_rate, paid=paid if i == 0 else 0,
                    mode=mode if i == 0 else "", note=note)
                written += 1
                bill += (ln["qty"] * ln_rate) if ln["type"] == "Issue" else 0
                if ln["type"] == "Issue":
                    given += ln["qty"]
                else:
                    got += ln["qty"]
        wb.save(XLSX)
        if typ == "Payment":
            flash("Saved ✓ %s • Payment ₹%s" % (cust, "{:,.0f}".format(paid)), "ok")
        else:
            flash("Saved ✓ %s • %d item line%s • Issued %s · Returned %s · Bill ₹%s" % (
                cust, written, "" if written == 1 else "s",
                ("%g" % given), ("%g" % got), "{:,.0f}".format(bill)), "ok")
        if f.get("again"):
            return redirect(back)
        if f.get("after") == "statement":
            return redirect(url_for("customer_detail", name=cust))
        return redirect(url_for("home"))
    pre = {"customer": request.args.get("customer", "").strip(),
        "type": request.args.get("type", "Issue").strip() or "Issue",
        "amount": request.args.get("amount", "").strip(),
        "date": request.args.get("date", "").strip(),
        "item": request.args.get("item", "").strip()}
    day = read_ledger()
    erows = sorted([r for r in day if r["customer"] == pre["customer"] and pre["customer"]],
        key=lambda x: (x["date"], x["row"]), reverse=True)[:6] if pre["customer"] else []
    if pre["type"] not in TYPES:
        pre["type"] = "Issue"
    return render_template("entry.html", items=items, custs=custs, modes=MODES,
        today=date.today().isoformat(), pre=pre, erows=erows)
@app.route("/customers")
def customers():
    rows = read_ledger(); items, custs = read_masters()
    s = summarize(rows, items, custs)
    q = request.args.get("q", "").strip()
    kh = s["khata"]
    only_due = request.args.get("due") == "1"
    if only_due:
        kh = [e for e in kh if e["pending"] > 0]
    return render_template("customers.html", khata=kh, all_khata=s["khata"], q=q, only_due=only_due)
@app.route("/customer/<name>")
def customer_detail(name):
    rows = read_ledger(); items, custs = read_masters()
    s = summarize(rows, items, custs)
    erows = sorted([r for r in rows if r["customer"] == name], key=lambda x: x["date"])
    info = next((e for e in s["khata"] if e["name"] == name),
        {"name": name, "billed": sum(r["bill"] for r in erows), "received": sum(r["paid"] for r in erows),
         "opening": 0, "given": 0, "returned": 0, "pending": 0, "outside": 0})
    if not info.get("pending"):
        info["pending"] = info["billed"] + info.get("opening", 0) - info["received"]
    master = next((c for c in custs if c["name"] == name), None)
    phone = next((c["phone"] for c in custs if c["name"] == name), "") or ""
    digits = "".join(ch for ch in phone if ch.isdigit())
    if len(digits) == 10:
        digits = "91" + digits
    wa = f"Namaskar {name}, DipSeva: Bill ₹{info['billed']:,.0f}, Jama ₹{info['received']:,.0f}, Baki ₹{info['pending']:,.0f}"
    unpaid = [r for r in erows if r["type"] == "Issue" and r["bill"]]
    bals = []
    bal = info.get("opening", 0) or 0
    for r in erows:
        bal += r["bill"] - r["paid"]
        bals.append(bal)
    last = erows[-1]["date_disp"] if erows else "—"
    return render_template("statement.html", info=info, rows=erows, master=master, wa=wa,
        phone=phone, wa_url=("https://wa.me/" + digits + "?text=" + quote(wa)) if digits else "",
        call_url=("tel:" + digits) if digits else "", lines=len(erows), bals=bals,
        first=erows[0]["date_disp"] if erows else "—", last=last,
        items_out=sum(r["qty"] or 0 for r in erows if r["type"] == "Issue") - sum(r["qty"] or 0 for r in erows if r["type"] == "Return"),
        bills=len(unpaid))
@app.route("/customer/<name>/phone", methods=["POST"])
def save_phone(name):
    digits = "".join(ch for ch in (request.form.get("phone") or "") if ch.isdigit() or ch == "+")
    wb = load_wb(); wc = wb["Customers"]
    hit = None
    for r in range(3, 203):
        if str(wc.cell(r, 1).value or "").strip() == name:
            hit = r
            break
    if hit is None:
        flash("Customer नाव sheet मध्ये सापडले नाही — spelling तपासा", "err")
        return redirect(url_for("customer_detail", name=name))
    if digits:
        wc.cell(hit, 3, digits)
        wb.save(XLSX)
        flash("Phone saved ✓ " + name, "ok")
    else:
        flash("नंबर रिकामा आहे", "err")
    return redirect(url_for("customer_detail", name=name))
@app.route("/calendar")
def calendar_view():
    rows = read_ledger()
    dm = day_map(rows)
    today = date.today()
    ds = request.args.get("d", "").strip()
    if not ds:
        ds = today.isoformat() if today.isoformat() in dm else max(list(dm.keys()) or [today.isoformat()])
    sel = parse_date(ds) or today
    ds = sel.isoformat()
    my, mm = sel.year, sel.month
    ms = request.args.get("m", "").strip()
    if len(ms) == 7 and ms[4] == "-":
        try:
            my, mm = int(ms[:4]), int(ms[5:])
        except Exception:
            my, mm = sel.year, sel.month
    weeks = []
    for wk in pycal.Calendar(0).monthdatescalendar(my, mm):
        cells = []
        for d in wk:
            iso = d.isoformat()
            e = dm.get(iso) or {}
            cells.append({"d": d.day, "iso": iso, "here": d.month == mm,
                "bill": e.get("bill", 0), "paid": e.get("paid", 0), "custn": e.get("custn", 0),
                "given": e.get("given", 0), "returned": e.get("returned", 0),
                "lines": e.get("lines", 0), "on": iso == ds, "is_today": iso == today.isoformat()})
        weeks.append(cells)
    groups = day_groups(rows, ds)
    st = month_stats(rows, my, mm)
    prevm = date(my, mm, 1) - timedelta(days=1)
    nextm = date(my + (1 if mm == 12 else 0), 1 if mm == 12 else mm + 1, 1)
    day = dm.get(ds) or {}
    return render_template("calendar.html", weeks=weeks, groups=groups, sel=ds,
        sel_disp=sel.strftime("%d-%b-%Y"), sel_day=sel.strftime("%A"), month_lbl=sel.strftime("%B %Y"),
        m="%04d-%02d" % (my, mm), prevm=prevm.strftime("%Y-%m"), nextm=nextm.strftime("%Y-%m"),
        st=st, day=day, today=today.isoformat(), wdays=["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"])


@app.route("/entry/<int:row>/edit", methods=["GET", "POST"])
def entry_edit(row):
    items, custs = read_masters()
    rec = next((r for r in read_ledger() if r["row"] == row), None)
    if rec is None:
        flash("Entry सापडली नाही — row %d रिकामी आहे" % row, "err")
        return redirect(url_for("home"))
    if request.method == "POST":
        f = request.form
        dt = parse_date(f.get("date")) or parse_date(rec["date"]) or date.today()
        cust = (f.get("customer") or rec["customer"]).strip()
        typ = (f.get("type") or rec["type"]).strip()
        if typ not in TYPES:
            typ = rec["type"]
        item = f.get("item", "").strip()
        qty = money(f.get("qty"), 0)
        rate = money(f.get("rate"), 0)
        amt = money(f.get("amount"), 0)
        paid = money(f.get("paid"), 0)
        if amt and qty:
            rate = round(amt / qty, 2)
        if not rate:
            rate = next((money(i["rate"], 0) for i in items if i["id"] == item), 0)
        mode = f.get("mode", "").strip()
        note = f.get("note", "").strip()
        if not cust:
            flash("Customer रिकामा आहे", "err")
            return redirect(url_for("entry_edit", row=row))
        if typ in ("Issue", "Return") and not item:
            flash("Item निवडा", "err")
            return redirect(url_for("entry_edit", row=row))
        if typ == "Payment" and paid <= 0:
            flash("Payment रक्कम लिहा", "err")
            return redirect(url_for("entry_edit", row=row))
        wb = load_wb(); ws = wb["Ledger"]
        write_line(ws, row, dt, cust, typ, item=item, qty=qty, rate=rate, paid=paid, mode=mode, note=note)
        wb.save(XLSX)
        flash("Updated ✓ %s • %s • row %d" % (cust, dt.strftime("%d-%b-%Y"), row), "ok")
        return redirect(url_for("customer_detail", name=cust))
    return render_template("edit.html", r=rec, items=items, custs=custs, modes=MODES, types=TYPES)


@app.route("/entry/<int:row>/delete", methods=["POST"])
def entry_delete(row):
    rec = next((r for r in read_ledger() if r["row"] == row), None)
    if rec is None:
        flash("Entry सापडली नाही", "err")
        return redirect(url_for("home"))
    wb = load_wb(); ws = wb["Ledger"]
    clear_row(ws, row)
    wb.save(XLSX)
    flash("Deleted ✓ %s • row %d" % (rec["customer"] or "—", row), "ok")
    if rec["customer"]:
        return redirect(url_for("customer_detail", name=rec["customer"]))
    return redirect(url_for("home"))


@app.route("/stock")
def stock():
    rows = read_ledger(); items, custs = read_masters()
    s = summarize(rows, items, custs)
    return render_template("stock.html", stock=s["stock"], low=s["low"])
@app.route("/api/summary")
def api_summary():
    rows = read_ledger(); items, custs = read_masters()
    s = summarize(rows, items, custs, request.args.get("today") or date.today().isoformat())
    return jsonify({"billed": s["billed"], "received": s["received"], "pending": s["pending"],
        "today_billed": s["today_billed"], "today_paid": s["today_paid"],
        "due_count": s["due_count"], "low_count": len(s["low"])})
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
