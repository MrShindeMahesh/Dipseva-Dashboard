from datetime import date
from flask import Flask, render_template, request, redirect, url_for, jsonify, flash
from store import XLSX, load_wb, parse_date, read_ledger, read_masters
from summary import summarize
app = Flask(__name__, template_folder="templates", static_folder="static")
app.secret_key = "dipseva-local-2026"
@app.route("/")
def home():
    today = request.args.get("today") or date.today().isoformat()
    rows = read_ledger(); items, custs = read_masters()
    s = summarize(rows, items, custs, today)
    recent = sorted([r for r in rows if r["date"]], key=lambda x: x["date"], reverse=True)[:15]
    return render_template("dashboard.html", s=s, items=items, custs=custs, recent=recent, today=today)
@app.route("/entry", methods=["GET", "POST"])
def entry():
    items, custs = read_masters()
    if request.method == "POST":
        f = request.form
        dt = parse_date(f.get("date")) or date.today()
        cust = f.get("customer", "").strip()
        typ = f.get("type", "Issue").strip()
        item = f.get("item", "").strip()
        try: qty = float(f.get("qty") or 0)
        except Exception: qty = 0
        rmap = {i["id"]: float(i["rate"] or 0) for i in items}
        try: rate = float(f.get("rate") or rmap.get(item, 0) or 0)
        except Exception: rate = rmap.get(item, 0)
        try: paid = float(f.get("paid") or 0)
        except Exception: paid = 0
        mode = f.get("mode", "").strip(); note = f.get("note", "").strip()
        if not cust:
            flash("Customer निवडा", "err"); return redirect(url_for("entry"))
        if typ in ("Issue", "Return") and not item:
            flash("Item निवडा", "err"); return redirect(url_for("entry"))
        wb = load_wb(); ws = wb["Ledger"]
        r = 3
        while r <= 1002 and not (ws.cell(r, 1).value in (None, "") and ws.cell(r, 2).value in (None, "") and ws.cell(r, 3).value in (None, "")):
            r += 1
        if r > 1002:
            flash("Ledger full", "err"); return redirect(url_for("entry"))
        ws.cell(r, 1, dt); ws.cell(r, 1).number_format = "DD-MMM-YYYY"
        ws.cell(r, 2, cust); ws.cell(r, 3, typ)
        if typ == "Payment":
            ws.cell(r, 9, paid); ws.cell(r, 10, mode or "Cash"); ws.cell(r, 11, note)
        else:
            ws.cell(r, 4, item); ws.cell(r, 6, qty); ws.cell(r, 7, rate)
            ws.cell(r, 5).value = f'=IF($D{r}="","",IFERROR(XLOOKUP($D{r},Items!$A:$A,Items!$B:$B),"ERR"))'
            ws.cell(r, 8).value = f'=IF($C{r}="Issue",IFERROR($F{r}*$G{r},0),0)'
            ws.cell(r, 9, paid if paid else None)
            ws.cell(r, 10, mode or ("Cash" if paid else None)); ws.cell(r, 11, note)
        wb.save(XLSX)
        flash("Saved ✓ " + cust + " • " + typ, "ok")
        return redirect(url_for("home"))
    return render_template("entry.html", items=items, custs=custs, today=date.today().isoformat())
@app.route("/customers")
def customers():
    rows = read_ledger(); items, custs = read_masters()
    s = summarize(rows, items, custs)
    q = request.args.get("q", "").strip().lower()
    kh = s["khata"]
    if q: kh = [e for e in kh if q in e["name"].lower()]
    return render_template("customers.html", khata=kh, q=q)
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
    wa = f"Namaskar {name}, DipSeva: Bill Rs{info['billed']:,.0f}, Jama Rs{info['received']:,.0f}, Baki Rs{info['pending']:,.0f}"
    return render_template("statement.html", info=info, rows=erows, master=master, wa=wa)
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
