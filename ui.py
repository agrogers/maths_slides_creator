from flask import Flask, render_template, request, redirect, url_for
import json
import os

app = Flask(__name__)

DATA_FILE = "question_data.json"
QUESTION_TYPES = [
    "add", "add3", "subtract", "multiply", "divide", "divide<11",
    "place_value", "place_value_reverse",
    "add_dec1", "add_fraction_same_denominator", "add_fraction_different_denominator",
    "perc10", "linear_equation_a*x=c", "linear_equation_a(x+b)=c", "linear_equation_ax+b=c"
]

def load_data():
    if os.path.exists(DATA_FILE):
        with open(DATA_FILE, "r") as f:
            return json.load(f)
    return {}

def save_data(data):
    with open(DATA_FILE, "w") as f:
        json.dump(data, f, indent=2)

def total_qty(question_sets):
    return sum(entry["qty"] for entries in question_sets.values() for entry in entries)

def flatten_entries(grade_data):
    flattened = []
    for qtype, entries in grade_data.items():
        for i, entry in enumerate(entries):
            flat_entry = entry.copy()
            flat_entry["qtype"] = qtype
            flat_entry["index"] = i
            flattened.append(flat_entry)
    return flattened

@app.route("/")
def index():
    data = load_data()
    grade_totals = {g: total_qty(qsets) for g, qsets in data.items()}

    # Optional sorting
    selected_grade = request.args.get("grade")
    sort_by = request.args.get("sort")

    selected_entries = []
    if selected_grade and selected_grade in data:
        selected_entries = flatten_entries(data[selected_grade])
        if sort_by in ("qtype", "tiers"):
            selected_entries.sort(key=lambda x: x.get(sort_by))

    return render_template("index.html",
                           data=data,
                           question_types=QUESTION_TYPES,
                           grade_totals=grade_totals,
                           selected_grade=selected_grade,
                           selected_entries=selected_entries,
                           sort_by=sort_by)

@app.route("/add", methods=["POST"])
def add():
    data = load_data()
    grade = request.form["grade"]
    qtype = request.form["qtype"]
    qty = int(request.form["qty"])
    min_val = json.loads(request.form["min"])
    max_val = json.loads(request.form["max"])
    tiers = json.loads(request.form["tiers"])
    fontsize = request.form.get("fontsize")

    entry = {
        "qty": qty,
        "min": min_val,
        "max": max_val,
        "tiers": tiers
    }
    if fontsize:
        entry["fontsize"] = int(fontsize)

    if grade not in data:
        data[grade] = {}
    if qtype not in data[grade]:
        data[grade][qtype] = []

    data[grade][qtype].append(entry)
    save_data(data)
    return redirect(url_for("index", grade=grade))

@app.route("/delete", methods=["POST"])
def delete():
    data = load_data()
    grade = request.form["grade"]
    qtype = request.form["qtype"]
    idx = int(request.form["index"])

    if grade in data and qtype in data[grade]:
        del data[grade][qtype][idx]
        if not data[grade][qtype]:
            del data[grade][qtype]
    save_data(data)
    return redirect(url_for("index", grade=grade))

@app.route("/edit", methods=["POST"])
def edit():
    data = load_data()
    grade = request.form["grade"]
    qtype = request.form["qtype"]
    idx = int(request.form["index"])

    if grade in data and qtype in data[grade] and 0 <= idx < len(data[grade][qtype]):
        entry = data[grade][qtype][idx]
        entry["qty"] = int(request.form["qty"])
        entry["min"] = json.loads(request.form["min"])
        entry["max"] = json.loads(request.form["max"])
        entry["tiers"] = json.loads(request.form["tiers"])
        fontsize = request.form.get("fontsize")
        if fontsize:
            entry["fontsize"] = int(fontsize)
        elif "fontsize" in entry:
            del entry["fontsize"]

    save_data(data)
    return redirect(url_for("index", grade=grade))

if __name__ == "__main__":
    app.run(debug=True)
