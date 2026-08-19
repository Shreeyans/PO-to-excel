from flask import Flask, request, render_template, send_file, jsonify
from pathlib import Path
import tempfile
import zipfile
import re
import uuid
import pdfplumber
import pandas as pd

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

ITEM_RE = re.compile(
    r"^(\d+)\s+(\d{8,14})\s*/\s+(.*?)\s+(\d{8})\s+([\d,]+)\s+PC\s+.*?([\d,]+\.\d{2})\s*$"
)
TCODE_RE = re.compile(r"^(T\d+)\s*(.*)$")

def first_value(pattern, text, default=""):
    m = re.search(pattern, text, re.I)
    return m.group(1).strip() if m else default

def parse_pdf(path):
    with pdfplumber.open(path) as pdf:
        lines = [
            x.strip()
            for p in pdf.pages
            for x in (p.extract_text() or "").splitlines()
            if x.strip()
        ]

    full = "\n".join(lines)

    vendor = first_value(
        r"Vendor Name\s*:\s*(.*?)\s+Delivery Site\s*:",
        full,
        "Unknown Vendor"
    )
    po_number = first_value(r"PO Number\s*:\s*(\d+)", full)
    category = first_value(r"Category Name\s*:\s*([^\n]+)", full)

    rows = []

    for i, line in enumerate(lines):
        m = ITEM_RE.match(line)
        if not m:
            continue

        sr = int(m.group(1))
        description_parts = [m.group(3).strip()]
        barcode = ""
        mrp = m.group(6).replace(",", "")

        # The requested Barcode is the T-code after the slash.
        if i + 1 < len(lines):
            tm = TCODE_RE.match(lines[i + 1])
            if tm:
                barcode = tm.group(1)
                remainder = tm.group(2).strip()
                if remainder:
                    description_parts.append(remainder)

        # Some supplied PDFs put MRP on a standalone line.
        if i + 2 < len(lines):
            candidate = lines[i + 2].replace(",", "")
            if re.fullmatch(r"\d+(?:\.\d{1,2})?", candidate):
                mrp = candidate

        description = re.sub(
            r"\s+",
            " ",
            " ".join(description_parts)
        ).strip()

        rows.append({
            "SR No": sr,
            "Barcode": barcode,
            "Product Description": description,
            "Qty": m.group(5).replace(",", ""),
            "PO Number": po_number,
            "MRP": mrp,
            "Category Name": category
        })

    return vendor, po_number, rows

def safe_name(value):
    value = re.sub(r'[\\/:*?"<>|]+', "", value)
    return value.strip() or "Unknown"

@app.route("/")
def index():
    return render_template("index.html")

@app.post("/api/process")
def process():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "Please select at least one PDF."}), 400

    work = Path(tempfile.mkdtemp(prefix="poexcel_"))
    outputs = []

    try:
        for uploaded in files:
            if not uploaded.filename.lower().endswith(".pdf"):
                continue

            pdf_path = work / uploaded.filename
            uploaded.save(pdf_path)

            vendor, po_number, rows = parse_pdf(pdf_path)

            df = pd.DataFrame(rows, columns=[
                "SR No",
                "Barcode",
                "Product Description",
                "Qty",
                "PO Number",
                "MRP",
                "Category Name"
            ])

            filename = f"{safe_name(vendor)}_{safe_name(po_number)}.xlsx"
            output = work / filename

            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="PO Data")
                ws = writer.book["PO Data"]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions

                widths = {
                    "A": 9, "B": 16, "C": 48, "D": 12,
                    "E": 16, "F": 12, "G": 26
                }
                for col, width in widths.items():
                    ws.column_dimensions[col].width = width

            outputs.append({
                "filename": filename,
                "path": str(output),
                "rows": len(df),
                "vendor": vendor,
                "po": po_number
            })

        if not outputs:
            return jsonify({"error": "No PDF files were selected."}), 400

        # One zip is convenient for multiple files, while individual files
        # remain available through the response metadata in this MVP.
        zip_path = work / "PO_Excel_Files.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for item in outputs:
                z.write(item["path"], item["filename"])

        return jsonify({
            "files": outputs,
            "zip": str(zip_path)
        })
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

@app.get("/api/download")
def download():
    path = Path(request.args.get("path", ""))
    if not path.is_file():
        return "File not found", 404
    return send_file(path, as_attachment=True, download_name=path.name)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080, debug=False)
