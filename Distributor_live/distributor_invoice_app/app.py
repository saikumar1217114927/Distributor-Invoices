"""
Distributor Commission-Payout Invoice Generator
================================================

Run:
    pip install -r requirements.txt
    python app.py

Then open http://127.0.0.1:5000 , upload the distributor Excel sheet,
fill in the payout month and invoice date, and download a ZIP containing
one Word (.docx) TAX INVOICE per distributor - built from the sample
CP Payout format, with GST auto-calculated per distributor:

  - No GST No on file                -> no tax lines, invoice = Total
  - GST No NOT starting with '33'    -> IGST @ 18%
  - GST No starting with '33'        -> SGST @ 9% + CGST @ 9%

The Excel 'Total' column is treated as the FINAL (GST-inclusive) payout
amount; GST is extracted back out of it (Total / 1.18 = base commission).

Expected Excel columns (case-insensitive, order doesn't matter):
    Partner Code | Partner Name | Total | GST | Address | PAN |
    Bank Name | Bank Account Number | ifsc code
('Address' and 'PAN' are optional - if missing, those lines are simply
omitted from the invoice. An 'RM' / RM-name column, if present, is
ignored.)
"""

import io
import os
import zipfile
import tempfile
import traceback
from datetime import datetime
from functools import wraps

import pandas as pd
from flask import Flask, request, render_template, send_file, flash, redirect, url_for, session

from invoice_engine import build_invoice_docx, safe_filename, compute_gst, is_blank_gst

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-me")

# Optional team password gate. Set the TEAM_PASSWORD environment variable
# on your host (e.g. Render -> Environment) to require a shared password
# before anyone can use the app. Leave it unset to allow open access.
TEAM_PASSWORD = os.environ.get("TEAM_PASSWORD")


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if TEAM_PASSWORD and not session.get("authed"):
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


@app.route("/login", methods=["GET", "POST"])
def login():
    if not TEAM_PASSWORD:
        return redirect(url_for("index"))
    if request.method == "POST":
        if request.form.get("password") == TEAM_PASSWORD:
            session["authed"] = True
            return redirect(request.args.get("next") or url_for("index"))
        flash("Incorrect password.")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.pop("authed", None)
    return redirect(url_for("login"))

# Map of accepted (lower-cased, stripped) header names -> canonical field
COLUMN_ALIASES = {
    "partner code": "partner_code",
    "partner name": "partner_name",
    "total": "total",
    "gst": "gst_no",
    "gst no": "gst_no",
    "address": "address",
    "pan": "pan_no",
    "pan no": "pan_no",
    "bank name": "bank_name",
    "bank account number": "account_number",
    "account number": "account_number",
    "ifsc code": "ifsc",
    "ifsc": "ifsc",
}

REQUIRED_FIELDS = ["partner_code", "partner_name", "total", "bank_name", "account_number", "ifsc"]


def format_month_period(raw_value):
    """HTML <input type=month> gives 'YYYY-MM' -> 'Apr-2023'."""
    dt = datetime.strptime(raw_value, "%Y-%m")
    return dt.strftime("%b-%Y")


def format_invoice_date(raw_value):
    """HTML <input type=date> gives 'YYYY-MM-DD' -> '30-Apr-2023'."""
    dt = datetime.strptime(raw_value, "%Y-%m-%d")
    return dt.strftime("%d-%b-%Y")


def load_distributors(file_storage):
    """Read the uploaded Excel file into a list of normalized dicts."""
    df = pd.read_excel(file_storage, dtype=str)  # read everything as text first
    # also grab a numeric-safe version for the Total column
    file_storage.seek(0)
    df_numeric = pd.read_excel(file_storage)

    normalized_cols = {}
    for col in df.columns:
        key = str(col).strip().lower()
        if key in COLUMN_ALIASES:
            normalized_cols[col] = COLUMN_ALIASES[key]
    df = df.rename(columns=normalized_cols)
    df_numeric = df_numeric.rename(columns=normalized_cols)

    missing = [f for f in REQUIRED_FIELDS if f not in df.columns]
    if missing:
        raise ValueError(
            "The uploaded sheet is missing required column(s): "
            + ", ".join(missing)
            + ". Expected headers like: Partner Code, Partner Name, Total, GST, "
              "Bank Name, Bank Account Number, ifsc code."
        )

    rows = []
    for i in range(len(df)):
        rows.append({
            "partner_code": str(df.iloc[i]["partner_code"]).strip(),
            "partner_name": str(df.iloc[i]["partner_name"]).strip(),
            "total": float(df_numeric.iloc[i]["total"]),
            "gst_no": df.iloc[i]["gst_no"] if "gst_no" in df.columns else None,
            "address": str(df.iloc[i]["address"]).strip() if "address" in df.columns and str(df.iloc[i]["address"]).strip().lower() != "nan" else "",
            "pan_no": str(df.iloc[i]["pan_no"]).strip() if "pan_no" in df.columns and str(df.iloc[i]["pan_no"]).strip().lower() != "nan" else "",
            "bank_name": str(df.iloc[i]["bank_name"]).strip(),
            "account_number": str(df.iloc[i]["account_number"]).strip(),
            "ifsc": str(df.iloc[i]["ifsc"]).strip(),
        })
    return rows


@app.route("/", methods=["GET"])
@login_required
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
@login_required
def generate():
    excel_file = request.files.get("excel_file")
    month_period_raw = (request.form.get("month_period") or "").strip()
    invoice_date_raw = (request.form.get("invoice_date") or "").strip()

    if not excel_file or excel_file.filename == "":
        flash("Please choose an Excel file to upload.")
        return redirect(url_for("index"))
    if not month_period_raw:
        flash("Please select the payout month/period.")
        return redirect(url_for("index"))
    if not invoice_date_raw:
        flash("Please select the invoice date.")
        return redirect(url_for("index"))

    try:
        month_period = format_month_period(month_period_raw)
    except ValueError:
        flash("Payout month/period is not a valid month.")
        return redirect(url_for("index"))

    try:
        invoice_date = format_invoice_date(invoice_date_raw)
    except ValueError:
        flash("Invoice date is not a valid date.")
        return redirect(url_for("index"))

    try:
        rows = load_distributors(excel_file)
    except Exception as exc:
        flash(f"Could not read the Excel file: {exc}")
        return redirect(url_for("index"))

    if not rows:
        flash("No distributor rows found in the uploaded sheet.")
        return redirect(url_for("index"))

    with tempfile.TemporaryDirectory() as tmp_dir:
        zip_buffer = io.BytesIO()
        summary_lines = ["Partner Code,Partner Name,Base,IGST,SGST,CGST,Total"]

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for row in rows:
                fname = f"{safe_filename(row['partner_code'])}_{safe_filename(month_period)}.docx"
                out_path = os.path.join(tmp_dir, fname)
                try:
                    gst = build_invoice_docx(
                        row=row,
                        month_period=month_period,
                        invoice_date=invoice_date,
                        ref_prefix=row["partner_code"],
                        out_path=out_path,
                    )
                except Exception:
                    traceback.print_exc()
                    continue
                zf.write(out_path, arcname=fname)
                summary_lines.append(
                    f"{row['partner_code']},{row['partner_name']},"
                    f"{gst['base']},{gst['igst']},{gst['sgst']},{gst['cgst']},{gst['total']}"
                )
            zf.writestr("summary.csv", "\n".join(summary_lines))

        zip_buffer.seek(0)

    zip_name = f"Distributor_Invoices_{safe_filename(month_period)}.zip"
    return send_file(
        zip_buffer,
        as_attachment=True,
        download_name=zip_name,
        mimetype="application/zip",
    )


if __name__ == "__main__":
    app.run(debug=True)
