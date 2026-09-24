"""Summary maths: single source for dashboard + khata + stock."""
from datetime import datetime, date, timedelta
def summarize(rows, items, custs, today=None):
    if today is None:
        today = date.today().isoformat()
    billed = sum(r["bill"] for r in rows)
    received = sum(r["paid"] for r in rows)
    opening_due = sum(float(c.get("opening_due") or 0) for c in custs)
    pending = billed + opening_due - received
    kh = {}
    for c in custs:
        kh[c["name"]] = {"name": c["name"], "billed": 0, "received": 0,
            "opening": float(c.get("opening_due") or 0), "given": 0, "returned": 0}
    for r in rows:
        name = r["customer"] or "CASH"
        e = kh.setdefault(name, {"name": name, "billed": 0, "received": 0, "opening": 0, "given": 0, "returned": 0})
        e["billed"] += r["bill"]; e["received"] += r["paid"]
        if r["type"] == "Issue":
            e["given"] += r["qty"] or 0
        if r["type"] == "Return":
            e["returned"] += r["qty"] or 0
    for e in kh.values():
        e["pending"] = e["billed"] + e["opening"] - e["received"]
        e["outside"] = (e["given"] or 0) - (e["returned"] or 0)
    khata = sorted([e for e in kh.values() if e["billed"] or e["received"] or e["opening"] or e["given"]],
        key=lambda x: x["pending"], reverse=True)
    st = {}
    for it in items:
        st[it["id"]] = {"id": it["id"], "name": it["name"], "marathi": it["marathi"],
            "opening": float(it["opening"] or 0), "given": 0, "returned": 0,
            "min": float(it["min"] or 0), "revenue": 0}
    for r in rows:
        if not r["item"]:
            continue
        e = st.setdefault(r["item"], {"id": r["item"], "name": r["item_name"], "marathi": "",
            "opening": 0, "given": 0, "returned": 0, "min": 0, "revenue": 0})
        if r["type"] == "Issue":
            e["given"] += r["qty"] or 0; e["revenue"] += r["bill"]
        if r["type"] == "Return":
            e["returned"] += r["qty"] or 0
    stock = list(st.values())
    for e in stock:
        e["inhand"] = e["opening"] + e["returned"] - e["given"]
        e["status"] = "LOW" if e["inhand"] < e["min"] else ("EMPTY" if e["inhand"] == 0 and (e["given"] or e["opening"]) else "OK")
    low = [e for e in stock if e["status"] == "LOW"]
    tb = sum(r["bill"] for r in rows if r["date"] == today)
    tp = sum(r["paid"] for r in rows if r["date"] == today)
    try: t = datetime.fromisoformat(today).date()
    except Exception: t = date.today()
    days = []
    for k in range(9, -1, -1):
        d = t - timedelta(days=k)
        iso = d.isoformat()
        days.append({"date": d.strftime("%d-%b"), "iso": iso,
            "billed": sum(r["bill"] for r in rows if r["date"] == iso),
            "received": sum(r["paid"] for r in rows if r["date"] == iso)})
    return {"billed": billed, "received": received, "pending": pending,
        "khata": khata, "stock": stock, "low": low,
        "today_billed": tb, "today_paid": tp, "days": days,
        "due_count": sum(1 for e in khata if e["pending"] > 0)}
