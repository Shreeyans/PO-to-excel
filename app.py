import io
import os
import re
import zipfile

import pdfplumber
from flask import Flask, request, render_template_string, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

COLUMNS = [
    "SR No",
    "Barcode",
    "Product Description",
    "Qty",
    "PO Number",
    "MRP",
    "Category Name",
]

HTML = """
<!doctype html>
<html>
<head>
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PO PDF → Excel</title>
<style>
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif;
margin:0;background:#f4f6fa;color:#172033}
.wrap{max-width:920px;margin:30px auto;padding:18px}
.card{background:#fff;border-radius:22px;padding:28px;box-shadow:0 4px 20px rgba(0,0,0,.05)}
h1{text-align:center;font-size:40px;margin:0 0 12px}
.sub{text-align:center;color:#667085;font-size:18px;line-height:1.45}
.drop{display:block;border:3px dashed #9aa8bf;border-radius:18px;padding:58px 15px;
text-align:center;margin-top:28px;cursor:pointer}
.drop b{font-size:22px}
input{display:none}
button{width:100%;padding:18px;margin-top:18px;border:0;border-radius:14px;
background:#111827;color:#fff;font-size:20px;font-weight:700}
.err{margin-top:16px;color:#c52828;white-space:pre-wrap;line-height:1.45}
.note{margin-top:20px;color:#667085;line-height:1.6}
</style>
</head>
<body>
<div class="wrap">
<div class="card">
<h1>PO PDF → Excel</h1>
<p class="sub">Upload one or more purchase-order PDFs and create a separate Excel file for each PDF.</p>

<form method="post" action="/convert" enctype="multipart/form-data">
<label class="drop" for="files">
<b>Tap here to select PO PDF files</b><br><br>
iPhone · iPad · Android · PC
</label>
<input id="files" name="files" type="file" accept=".pdf,application/pdf" multiple>
<button type="submit">Convert to Excel</button>
</form>

{% if errors %}
<div class="err">{{ errors|join("\\n") }}</div>
{% endif %}

<div class="note">
<b>Excel columns:</b><br>
SR No · Barcode · Product Description · Qty · PO Number · MRP · Category Name
<br><br>
<b>Barcode:</b> the T-code immediately after “/”.
<br>
The converter handles split PDF lines, page breaks, optional MRP columns,
and MRP values embedded in descriptions.
</div>
</div>
</div>
</body>
</html>
"""

def clean(value):
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()

def number(value):
    m = re.search(r"-?\d[\d,]*(?:\.\d+)?", clean(value))
    return m.group(0).replace(",", "") if m else ""

def extract_tcode(value):
    s = clean(value)

    # IMPORTANT: do not accept words such as TOP as barcodes.
    m = re.search(r"/\s*(T\d{4,})\b", s, re.I)
    if m:
        return m.group(1).upper()

    m = re.search(r"\b(T\d{4,})\b", s, re.I)
    return m.group(1).upper() if m else ""

def parse_header(text):
    vendor = ""
    po = ""
    category = ""

    m = re.search(
        r"Vendor Name\s*:\s*(.*?)\s+Delivery Site",
        text,
        re.I
    )
    if m:
        vendor = clean(m.group(1))

    m = re.search(
        r"PO Number\s*:\s*([A-Za-z0-9_-]+)",
        text,
        re.I
    )
    if m:
        po = m.group(1).strip()

    m = re.search(
        r"Category Name\s*:\s*([^\n]+)",
        text,
        re.I
    )
    if m:
        category = clean(m.group(1))

    return vendor, po, category

def derive_mrp_from_description(description):
    """
    Some PO templates have no MRP column.
    In those files the MRP appears inside the product description,
    usually as the final 3-5 digit number.

    PDF extraction sometimes splits digits, e.g. 144\\n9.
    """
    d = clean(description)

    # Join spaces/newlines between digits: 144 9 -> 1449.
    d = re.sub(r"(?<=\d)\s+(?=\d)", "", d)

    candidates = re.findall(
        r"(?<!\d)(\d{3,5}(?:\.\d{1,2})?)(?!\d)",
        d
    )

    if not candidates:
        return "", d

    mrp = candidates[-1]

    # Remove only the selected MRP token from the description.
    cleaned = re.sub(
        r"(?<!\d)" + re.escape(mrp) + r"(?!\d)",
        "",
        d,
        count=1
    )

    cleaned = re.sub(r"\s+([,])", r"\1", cleaned)
    cleaned = re.sub(r"[,]\s*[,]", ",", cleaned)
    cleaned = cleaned.strip(" ,_-")

    return mrp, cleaned

def group_key(barcode, description):
    """
    Used only for recovering a digit lost by PDF extraction.
    Many size variants in the same product family share the first
    five digits after T.
    """
    b = clean(barcode).upper()
    prefix = b[:6] if b.startswith("T") and len(b) >= 6 else b

    d = clean(description).lower()
    d = re.sub(
        r"\b(?:xs|s|m|l|xl|2xl|3xl|4xl|5xl|6xl|7xl|8xl)\b",
        "",
        d
    )
    d = re.sub(r"\d+", "", d)
    d = re.sub(r"\s+", " ", d).strip(" ,_-")

    return prefix + "|" + d[:90]


def extract_item_rows(pdf):
    """
    Page-by-page PO table extraction with a cross-page continuation state.

    The supplied Jeyachandran PDFs vary in:
      - 20/21/22/23 table columns
      - optional MRP column
      - wrapped descriptions
      - page-break splits
      - serial numbers split as "10" + "0" after 99

    We therefore identify item rows from the stable Item/Bar Code,
    T-code and Qty fields and enumerate rows ourselves.
    """
    raw = []
    pending_index = None

    for page in pdf.pages:
        tables = page.extract_tables()

        for table in tables:
            if not table:
                continue

            header_index = None

            for i, row in enumerate(table[:10]):
                joined = " ".join(
                    clean(c).lower()
                    for c in (row or [])
                    if c
                )

                if (
                    "sl" in joined
                    and "item/bar" in joined
                    and "description" in joined
                    and "qty" in joined
                ):
                    header_index = i
                    break

            if header_index is None:
                continue

            header = table[header_index]
            normalized_header = [
                clean(c).lower() if c else ""
                for c in header
            ]

            def find_column(text):
                for idx, value in enumerate(normalized_header):
                    if text in value:
                        return idx
                return None

            barcode_idx = find_column("item/bar")
            description_idx = find_column("description")
            qty_idx = find_column("qty")
            mrp_idx = find_column("mrp")
            has_mrp_column = mrp_idx is not None

            for row in table[header_index + 1:]:
                cells = [clean(c) for c in (row or [])]

                if not cells:
                    continue

                row_text = " ".join(cells)

                if re.search(
                    r"^(sub total|total value|taxable value|taxvalue|"
                    r"total net value|terms\s*&\s*conditions)\b",
                    row_text,
                    re.I
                ):
                    continue

                item_cell = (
                    cells[barcode_idx]
                    if barcode_idx is not None
                    and barcode_idx < len(cells)
                    else ""
                )

                description = (
                    cells[description_idx]
                    if description_idx is not None
                    and description_idx < len(cells)
                    else ""
                )

                qty = (
                    number(cells[qty_idx])
                    if qty_idx is not None
                    and qty_idx < len(cells)
                    else ""
                )

                mrp = (
                    number(cells[mrp_idx])
                    if mrp_idx is not None
                    and mrp_idx < len(cells)
                    else ""
                )

                barcode = extract_tcode(item_cell)

                # Full item row.
                if qty:
                    raw.append({
                        "barcode": barcode,
                        "description": description,
                        "qty": qty,
                        "mrp": mrp,
                        "explicit_mrp": has_mrp_column,
                    })

                    # If T-code was pushed to the next page,
                    # keep this row pending globally across pages/tables.
                    pending_index = (
                        len(raw) - 1
                        if not barcode
                        else None
                    )
                    continue

                # Continuation row. In the supplied PDFs this is usually
                # a T-code + description with no Qty because the previous
                # item's cell was split at a page boundary.
                if barcode:
                    if (
                        pending_index is not None
                        and pending_index < len(raw)
                        and not raw[pending_index]["barcode"]
                    ):
                        raw[pending_index]["barcode"] = barcode

                        if description:
                            raw[pending_index]["description"] = clean(
                                raw[pending_index]["description"]
                                + " "
                                + description
                            )

                        if mrp:
                            raw[pending_index]["mrp"] = mrp
                            raw[pending_index]["explicit_mrp"] = has_mrp_column

                        pending_index = None

    raw = [
        r for r in raw
        if r["barcode"] and r["qty"]
    ]

    # Derive MRP for templates with no dedicated MRP column.
    for record in raw:
        if not record["explicit_mrp"]:
            mrp, desc = derive_mrp_from_description(
                record["description"]
            )
            record["mrp"] = mrp
            record["description"] = desc
        else:
            record["description"] = clean(record["description"])

    # Recover a digit lost by PDF extraction using the same T-code family.
    def family_key(record):
        barcode = clean(record["barcode"]).upper()
        prefix = barcode[:6] if len(barcode) >= 6 else barcode

        desc = clean(record["description"]).lower()
        desc = re.sub(
            r"\b(?:xs|s|m|l|xl|2xl|3xl|4xl|5xl|6xl|7xl|8xl)\b",
            "",
            desc
        )
        desc = re.sub(r"\d+", "", desc)
        desc = re.sub(r"\s+", " ", desc).strip(" ,_-")

        return prefix + "|" + desc[:90]

    family_values = {}

    for record in raw:
        if not record["explicit_mrp"]:
            if re.fullmatch(
                r"\d{3,5}(?:\.\d{1,2})?",
                record["mrp"] or ""
            ):
                family_values.setdefault(
                    family_key(record),
                    []
                ).append(record["mrp"])

    for record in raw:
        if record["explicit_mrp"]:
            # Explicit 0.00 is valid and must not be overwritten.
            continue

        candidates = family_values.get(
            family_key(record),
            []
        )

        if not candidates:
            continue

        dominant = max(
            set(candidates),
            key=candidates.count
        )

        current = record["mrp"] or ""

        if (
            not current
            or (
                current.isdigit()
                and len(current) < len(dominant.split(".")[0])
                and dominant.startswith(current)
            )
        ):
            record["mrp"] = dominant

            # Remove a truncated MRP fragment left at the end
            # of the product description, e.g. ",44" from ",449".
            record["description"] = re.sub(
                r"[, ]*\d{1,4}\s*$",
                "",
                record["description"]
            ).strip(" ,_-")

    final_rows = []

    for sr_no, record in enumerate(raw, start=1):
        final_rows.append({
            "SR No": sr_no,
            "Barcode": record["barcode"],
            "Product Description": clean(record["description"]),
            "Qty": record["qty"],
            "MRP": record["mrp"],
        })

    return final_rows

def parse_pdf(data):
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        first_page_text = pdf.pages[0].extract_text() or ""

        vendor, po_number, category = parse_header(
            first_page_text
        )

        rows = extract_item_rows(pdf)

        if not rows:
            raise ValueError(
                "No PO item rows could be detected in this PDF."
            )

        # Fail closed rather than producing a potentially wrong Excel.
        incomplete = [
            r for r in rows
            if not r["Barcode"]
            or not r["Qty"]
            or not r["MRP"]
        ]

        if incomplete:
            raise ValueError(
                f"PO rows were found, but {len(incomplete)} rows "
                "are missing Barcode, Qty or MRP. The file was not "
                "converted so that incorrect data is not produced."
            )

        return vendor, po_number, category, rows

def safe_filename(value):
    value = clean(value) or "Unknown"
    value = re.sub(
        r'[\\/:*?"<>|]+',
        "_",
        value
    )
    return value.strip(" ._")

def create_excel(rows, po_number, category):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "PO Items"

    sheet.append(COLUMNS)

    for cell in sheet[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(
            horizontal="center"
        )

    for row in rows:
        sheet.append([
            row["SR No"],
            row["Barcode"],
            row["Product Description"],
            row["Qty"],
            po_number,
            row["MRP"],
            category,
        ])

    sheet.freeze_panes = "A2"

    widths = {
        "A": 10,
        "B": 18,
        "C": 55,
        "D": 12,
        "E": 18,
        "F": 14,
        "G": 28,
    }

    for column, width in widths.items():
        sheet.column_dimensions[column].width = width

    for row in sheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)

    return output.getvalue()

@app.get("/")
def home():
    return render_template_string(HTML)

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/convert")
def convert():
    files = [
        f for f in request.files.getlist("files")
        if f and f.filename.lower().endswith(".pdf")
    ]

    if not files:
        return render_template_string(
            HTML,
            errors=["Please select at least one PDF."]
        )

    generated = []
    errors = []

    for uploaded in files:
        try:
            data = uploaded.read()

            vendor, po_number, category, rows = parse_pdf(
                data
            )

            filename = (
                f"{safe_filename(vendor)}_"
                f"{safe_filename(po_number)}.xlsx"
            )

            generated.append(
                (
                    filename,
                    create_excel(
                        rows,
                        po_number,
                        category
                    )
                )
            )

        except Exception as exc:
            errors.append(
                f"{uploaded.filename}: {exc}"
            )

    if not generated:
        return render_template_string(
            HTML,
            errors=errors
        ), 422

    # One PDF -> one Excel file.
    if len(generated) == 1 and not errors:
        filename, data = generated[0]

        return send_file(
            io.BytesIO(data),
            as_attachment=True,
            download_name=filename,
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    # Multiple PDFs -> ZIP with separate Excel files.
    archive = io.BytesIO()

    with zipfile.ZipFile(
        archive,
        "w",
        zipfile.ZIP_DEFLATED
    ) as z:
        for filename, data in generated:
            z.writestr(filename, data)

        if errors:
            z.writestr(
                "conversion_errors.txt",
                "\n".join(errors)
            )

    archive.seek(0)

    return send_file(
        archive,
        as_attachment=True,
        download_name="PO_Excel_Files.zip",
        mimetype="application/zip"
    )

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", "10000"))
    )
