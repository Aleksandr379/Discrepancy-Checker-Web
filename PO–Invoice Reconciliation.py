from flask import Flask, request, render_template_string, send_from_directory, abort
import pandas as pd
import pdfplumber
import os
import uuid
from werkzeug.utils import secure_filename

app = Flask(__name__)

# ---------------- CONFIG ----------------
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"pdf", "xlsx", "xls"}

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024


# ---------------- HTML ----------------
HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>PO-Invoice Reconciliation</title>
</head>
<body>

<h2>PO vs Invoice Reconciliation System</h2>

<form method="POST" enctype="multipart/form-data">
    <label>Purchase Order:</label>
    <input type="file" name="po_file" required><br><br>

    <label>Invoice:</label>
    <input type="file" name="invoice_file" required><br><br>

    <button type="submit">Process</button>
</form>

{% if error %}
<p style="color:red;"><b>{{ error }}</b></p>
{% endif %}

{% if report_id %}
<h3>Reports Generated</h3>
<ul>
    <li><a href="/download/{{report_id}}/excel">Download Excel</a></li>
    <li><a href="/download/{{report_id}}/pdf">Download PDF</a></li>
</ul>
{% endif %}

</body>
</html>
"""


# ---------------- HELPERS ----------------

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def safe_float(v):
    try:
        return float(str(v).replace(",", "").strip())
    except:
        return 0.0


def normalize_values(v):
    if pd.isna(v):
        return None
    if isinstance(v, str):
        return v.strip().lower()
    try:
        return round(float(v), 2)
    except:
        return str(v).strip().lower()


# ---------------- COLUMN NORMALIZATION (ROBUST) ----------------

def normalize_columns(df):
    df = df.copy()
    df.columns = [str(c).strip().lower() for c in df.columns]

    mapping = {
        "po number": ["po number", "po", "po id", "po no", "po#", "purchase order"],
        "invoice number": ["invoice number", "invoice id", "inv no", "invoice"],
        "vendor": ["vendor", "supplier"],
        "quantity": ["qty", "quantity"],
        "amount": ["amount", "total", "total amount"],
        "currency": ["currency"]
    }

    rename_map = {}

    for std, variants in mapping.items():
        for col in df.columns:
            col_clean = col.replace(" ", "")
            if any(v.replace(" ", "") in col_clean for v in variants):
                rename_map[col] = std
                break

    df.rename(columns=rename_map, inplace=True)
    return df


# ---------------- FILE READERS (FIXED PDF LOGIC) ----------------

def read_excel(path):
    df = pd.read_excel(path)
    if df.empty:
        raise ValueError("Excel file is empty")
    return df


def read_pdf(path):
    rows = []

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    for row in table:
                        if row and any(cell is not None for cell in row):
                            rows.append(row)

    if not rows:
        raise ValueError("No tables found in PDF")

    headers = rows[0]
    data = rows[1:]

    return pd.DataFrame(data, columns=headers)


def extract_file(path):
    if path.endswith(".pdf"):
        return read_pdf(path)
    return read_excel(path)


# ---------------- RECONCILIATION (IMPROVED) ----------------

def reconcile(po_df, inv_df):

    po_df = normalize_columns(po_df)
    inv_df = normalize_columns(inv_df)

    if "po number" not in po_df.columns:
        raise ValueError("PO file missing 'PO Number'")
    if "po number" not in inv_df.columns:
        raise ValueError("Invoice file missing 'PO Number'")

    results = []

    total_po = 0
    total_inv = 0
    discrepancies = 0

    for _, po in po_df.iterrows():

        po_number = po.get("po number")
        if pd.isna(po_number):
            continue

        match = inv_df[
            inv_df["po number"].astype(str).str.strip()
            == str(po_number).strip()
        ]

        po_amount = safe_float(po.get("amount", 0))
        total_po += po_amount

        if match.empty:
            results.append({
                "PO Number": po_number,
                "Vendor": po.get("vendor"),
                "Field": "ALL",
                "PO Value": po_amount,
                "Invoice Value": "MISSING",
                "Variance": po_amount,
                "Status": "MISSING INVOICE"
            })
            discrepancies += 1
            continue

        inv = match.iloc[0]

        inv_amount = safe_float(inv.get("amount", 0))
        total_inv += inv_amount

        fields = ["vendor", "quantity", "amount", "currency"]

        row_issue = False

        for f in fields:

            po_val = normalize_values(po.get(f))
            inv_val = normalize_values(inv.get(f))

            if po_val != inv_val:

                variance = 0
                if f == "amount":
                    variance = inv_amount - po_amount

                # SAFE FINANCE LOGIC (percentage-based)
                if po_amount and abs(variance) / po_amount > 0.1:
                    status = "MAJOR VARIANCE"
                else:
                    status = "MINOR VARIANCE"

                results.append({
                    "PO Number": po_number,
                    "Vendor": po.get("vendor"),
                    "Field": f,
                    "PO Value": po.get(f),
                    "Invoice Value": inv.get(f),
                    "Variance": variance,
                    "Status": status
                })

                row_issue = True
                discrepancies += 1

        if not row_issue:
            results.append({
                "PO Number": po_number,
                "Vendor": po.get("vendor"),
                "Field": "ALL",
                "PO Value": po_amount,
                "Invoice Value": inv_amount,
                "Variance": 0,
                "Status": "MATCH"
            })

    total_rows = len(po_df)

    match_rate = round(
        100 - (discrepancies / total_rows * 100),
        2
    ) if total_rows else 0

    summary = {
        "Total PO Value": total_po,
        "Total Invoice Value": total_inv,
        "Total Variance": total_inv - total_po,
        "Discrepancies": discrepancies,
        "Match Rate (%)": match_rate
    }

    return pd.DataFrame(results), summary


# ---------------- REPORTS ----------------

def create_excel(df, summary, path):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame([summary]).to_excel(writer, sheet_name="Summary", index=False)
        df.to_excel(writer, sheet_name="Details", index=False)


def create_pdf(df, summary, path):

    from reportlab.pdfgen import canvas

    c = canvas.Canvas(path)
    y = 800

    c.drawString(50, y, "PO-INVOICE RECONCILIATION REPORT")
    y -= 40

    for k, v in summary.items():
        c.drawString(50, y, f"{k}: {v}")
        y -= 20

    y -= 20
    c.drawString(50, y, "DETAILS (Top 25):")
    y -= 20

    for _, row in df.head(25).iterrows():
        line = f"{row['PO Number']} | {row['Field']} | {row['Status']}"
        c.drawString(50, y, line[:90])
        y -= 15

        if y < 50:
            c.showPage()
            y = 800

    c.save()


# ---------------- ROUTES ----------------

@app.route("/", methods=["GET", "POST"])
def home():

    error = None
    report_id = None

    try:
        if request.method == "POST":

            po = request.files.get("po_file")
            inv = request.files.get("invoice_file")

            if not po or not inv:
                raise ValueError("Files missing")

            if not allowed_file(po.filename) or not allowed_file(inv.filename):
                raise ValueError("Only PDF or Excel files allowed")

            po_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}_{secure_filename(po.filename)}")
            inv_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4()}_{secure_filename(inv.filename)}")

            po.save(po_path)
            inv.save(inv_path)

            po_df = extract_file(po_path)
            inv_df = extract_file(inv_path)

            report_df, summary = reconcile(po_df, inv_df)

            report_id = str(uuid.uuid4())

            excel_path = os.path.join(UPLOAD_FOLDER, f"{report_id}.xlsx")
            pdf_path = os.path.join(UPLOAD_FOLDER, f"{report_id}.pdf")

            create_excel(report_df, summary, excel_path)
            create_pdf(report_df, summary, pdf_path)

    except Exception as e:
        error = str(e)

    return render_template_string(HTML, error=error, report_id=report_id)


@app.route("/download/<report_id>/<filetype>")
def download(report_id, filetype):

    if filetype not in ["excel", "pdf"]:
        abort(404)

    filename = f"{report_id}.xlsx" if filetype == "excel" else f"{report_id}.pdf"

    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)


# ---------------- RUN ----------------

if __name__ == "__main__":
    app.run(debug=True)
