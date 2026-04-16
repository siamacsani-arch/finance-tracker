#!/usr/bin/env python3
"""
app.py — Finance Dashboard Flask server
Run: python3 app.py
"""
from flask import Flask, jsonify, request, render_template
from pathlib import Path
import db

BASE = Path(__file__).parent
app = Flask(__name__, template_folder=str(BASE / "templates"), static_folder=str(BASE / "static"))


# ── Startup: init DB and auto-import if empty ──────────────────────────────────
def startup():
    db.init_db()
    db.populate_budgets()
    conn = db.get_conn()
    count = conn.execute("SELECT COUNT(*) FROM transactions").fetchone()[0]
    conn.close()
    if count == 0:
        xlsx = BASE / "Financial Tracker.xlsx"
        if xlsx.exists():
            print("DB is empty — importing from Excel...")
            db.import_from_excel(str(xlsx))
        else:
            print("Warning: Financial Tracker.xlsx not found, DB will be empty.")
    else:
        print(f"DB has {count} transactions — skipping auto-import.")


# ── Routes ─────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/summary")
def api_summary():
    year = int(request.args.get("year", 2026))
    month = int(request.args.get("month", 4))
    return jsonify(db.get_summary(year, month))


@app.route("/api/budget_vs_actual")
def api_budget_vs_actual():
    year = int(request.args.get("year", 2026))
    month = int(request.args.get("month", 4))
    return jsonify(db.get_budget_vs_actual(year, month))


@app.route("/api/transactions")
def api_transactions():
    year = int(request.args.get("year", 2026))
    month = int(request.args.get("month", 4))
    txn_type = request.args.get("type", "")
    category = request.args.get("category", "")
    return jsonify(db.get_transactions(
        year, month,
        txn_type=txn_type if txn_type else None,
        category=category if category else None,
    ))


@app.route("/api/trends")
def api_trends():
    year = int(request.args.get("year", 2026))
    return jsonify(db.get_trends(year))


@app.route("/api/spending_by_category")
def api_spending_by_category():
    year = int(request.args.get("year", 2026))
    month = int(request.args.get("month", 4))
    return jsonify(db.get_spending_by_category(year, month))


@app.route("/api/budget", methods=["PUT"])
def api_update_budget():
    data = request.get_json()
    if not data or "category" not in data or "amount" not in data:
        return jsonify({"error": "category and amount required"}), 400
    try:
        db.upsert_budget(data["category"], float(data["amount"]))
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/transaction/<int:txn_id>/category", methods=["PUT"])
def api_update_txn_category(txn_id):
    data = request.get_json()
    if not data or "category" not in data:
        return jsonify({"error": "category required"}), 400
    db.update_transaction_category(txn_id, data["category"])
    return jsonify({"ok": True})


@app.route("/api/categories")
def api_categories():
    """Return distinct categories for filter dropdowns."""
    conn = db.get_conn()
    rows = conn.execute("SELECT DISTINCT category FROM transactions ORDER BY category").fetchall()
    conn.close()
    return jsonify([r["category"] for r in rows])


if __name__ == "__main__":
    startup()
    app.run(host="0.0.0.0", port=5001, debug=False)
