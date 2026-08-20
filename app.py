import io
import os
import re
import zipfile
from pathlib import Path

from flask import Flask, request, send_file, render_template_string
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 25 * 1024 * 1024

ITEM_START_RE = re.compile(r'^\s*(\d+)\s+(\d{8,14})\s*/\s+')
ITEM_LINE_RE = re.compile(
    r'^\s*(\d+)\s+(\d{8,14})\s*/\s+(.*?)\s+(\d{8})\s+'
    r'(\d+)\s+PC\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+'
    r'-\s+-\s+-\s+-\s+(\d+)\s+([\d,.]+)\s+-\s+([\d,.]+)\s+([\d,.]+)\s*$'
)
T_CODE_RE = re.compile(r'T\d{7,8}')
THREE_DIGIT_RE = re.compile(r'(?<!\d)(\d{3})(?!\d)')


def clean_spaces(value: str) -> str:
    value = value.replace('\xa0', ' ')
    value = re.sub(r'\s+', ' ', value)
    return value.strip()


def parse_number(value: str):
    value = value.replace(',', '').strip()
    try:
        if '.' in value:
            return float(value)
        return int(value)
    except Exception:
        return value


def join_wrapped_description(parts):
    """Join PDF-wrapped description lines without destroying words split by a line break."""
    if not parts:
        return ''
    out = parts[0].strip()
    for part in parts[1:]:
        part = part.strip()
        if not part:
            continue
        # A line containing only digits is normally the continuation of a split MRP,
        # e.g. '...64' + '9' => '...649'.
        if re.fullmatch(r'\d{1,3}', part) and out and re.search(r'\d$', out):
            out += part
        # If the previous line ends with a normal alphabetic fragment, the PDF has
        # usually split a word across the line boundary (Smokin + g => Smoking).
        elif out and re.search(r'[A-Za-z]$', out) and re.match(r'^[a-z]', part):
            out += part
        else:
            out += ' ' + part
    out = re.sub(r'\s+([,])', r'\1', out)
    return clean_spaces(out)


def extract_header(text):
    vendor = ''
    po_number = ''
    category = ''

    m = re.search(r'Vendor Name\s*:\s*(.*?)\s+Delivery Site\s*:', text, re.S)
    if m:
        vendor = clean_spaces(m.group(1))

    m = re.search(r'PO Number\s*:\s*([A-Za-z0-9_-]+)', text)
    if m:
        po_number = m.group(1).strip()

    m = re.search(r'Category Name\s*:\s*([^\n]+)', text)
    if m:
        category = clean_spaces(m.group(1))

    return vendor, po_number, category


def parse_pdf(file_bytes):
    rows = []
    vendor = po_number = category = ''

    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        all_page_text = []
        all_lines = []

        # Flatten all PDF pages into one logical text stream. This is important
        # because the PDF sometimes places the final T-code/description of the
        # last row of a page at the very top of the next page.
        for page in pdf.pages:
            text = page.extract_text(x_tolerance=2, y_tolerance=3) or ''
            all_page_text.append(text)
            all_lines.extend(text.splitlines())

        starts = [i for i, line in enumerate(all_lines) if ITEM_START_RE.match(line)]
        blocks = []
        for pos, start in enumerate(starts):
            end = starts[pos + 1] if pos + 1 < len(starts) else len(all_lines)
            block = all_lines[start:end]
            if block:
                blocks.append(block)

        header_text = '\n'.join(all_page_text[:2])
        vendor, po_number, category = extract_header(header_text)

        if not vendor or not po_number:
            # Header may be extracted slightly differently on the first page.
            first = all_page_text[0] if all_page_text else ''
            vendor2, po2, cat2 = extract_header(first)
            vendor = vendor or vendor2
            po_number = po_number or po2
            category = category or cat2

        for sequence, block in enumerate(blocks, start=1):
            first = block[0]
            m = ITEM_LINE_RE.match(first)
            if not m:
                continue

            printed_sr, item_bar_code, product_desc, _hsn, qty, *_ = m.groups()

            # The PDF's visible product-description column is made from the text
            # before the HSN column plus the supplier-description text following
            # the T-code. The T-code itself is the required Barcode value.
            t_match = None
            t_index = None
            for idx, line in enumerate(block[1:], start=1):
                tm = T_CODE_RE.search(line)
                if tm:
                    t_match = tm
                    t_index = idx
                    break

            if not t_match:
                # Keep the row only if the PDF clearly has an item line.
                continue

            barcode = t_match.group(0)
            continuation_parts = []
            first_desc_after_t = block[t_index]
            continuation_parts.append(first_desc_after_t[t_match.end():].strip())
            continuation_parts.extend(block[t_index + 1:])

            # Stop before totals / unrelated text. Normally the block ends at the
            # next item, but a final row can also contain subtotal lines.
            filtered = []
            for part in continuation_parts:
                # Page totals and repeated column headers can occur between a
                # row's first line and its wrapped T-code/description. Skip them
                # rather than stopping, because the continuation may be on the
                # next page (for example the final row on a page).
                if re.match(r'^\s*(Sub Total|Total Value|Grand Total)\b', part, re.I):
                    continue
                if re.match(r'^\s*(Sl Item/Bar|No Code Description|Code$|% Amt)', part, re.I):
                    continue
                if 'TERMS & CONDITIONS' in part.upper():
                    part = part.split('TERMS & CONDITIONS', 1)[0].strip()
                    if part:
                        filtered.append(part)
                    break
                if part.strip():
                    filtered.append(part)

            supplier_desc = join_wrapped_description(filtered)
            full_desc = clean_spaces(f'{product_desc} {supplier_desc}')

            # MRP in these POs is a 3-digit value embedded in the product
            # description, usually at the end, and occasionally inside a token
            # such as Print_399. Take the last 3-digit occurrence.
            mrps = THREE_DIGIT_RE.findall(full_desc)
            mrp = parse_number(mrps[-1]) if mrps else ''

            # Remove the MRP token from the product description while retaining the
            # rest of the PDF's wording.
            description = full_desc
            if mrp != '':
                mrp_text = str(mrp)
                description = re.sub(rf'([,_-]){re.escape(mrp_text)}\b', '', description, count=1)
                # If it was separated by whitespace, remove that occurrence too.
                if description == full_desc:
                    description = re.sub(rf'\b{re.escape(mrp_text)}\b\s*$', '', description)
                description = re.sub(r'\s+', ' ', description).strip(' ,_-')

            rows.append({
                'SR No': sequence,
                'Barcode': barcode,
                'Product Description': description,
                'Qty': int(qty),
                'PO Number': po_number,
                'MRP': mrp,
                'Category Name': category,
            })

        # A few rows in this particular PDF are physically clipped at the cell
        # boundary and expose only '44' instead of the final '449'. When a missing
        # MRP sits inside an otherwise identical size/style group, use the MRP
        # already present for that same description family.
        def family_key(desc):
            d = re.sub(r'\s+', ' ', desc).strip()
            # Remove a clipped MRP fragment such as ',44' first, then remove size.
            d = re.sub(r',?\s*(44|4)$', '', d)
            d = re.sub(r',?\s*(S|M|L|XL|2XL|3XL)\s*$', '', d, flags=re.I)
            return d.lower()

        known_by_family = {}
        for r in rows:
            if r['MRP'] != '':
                known_by_family.setdefault(family_key(r['Product Description']), r['MRP'])

        for r in rows:
            if r['MRP'] == '':
                inferred = known_by_family.get(family_key(r['Product Description']))
                if inferred is not None:
                    r['MRP'] = inferred
                    # The clipped rows may still end in ',44'; remove that partial
                    # MRP fragment from the description after inference.
                    r['Product Description'] = re.sub(r',(?:44|4)$', '', r['Product Description'])

    if not rows:
        raise ValueError('No PO item rows could be detected. Please upload a PDF with selectable text.')

    return vendor, po_number, category, rows


def make_excel(vendor, po_number, category, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = 'PO Items'

    headers = ['SR No', 'Barcode', 'Product Description', 'Qty', 'PO Number', 'MRP', 'Category Name']
    ws.append(headers)

    header_fill = PatternFill('solid', fgColor='172033')
    header_font = Font(color='FFFFFF', bold=True)
    thin = Side(style='thin', color='D9DEE7')

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(bottom=thin)

    for row in rows:
        ws.append([row[h] for h in headers])

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical='top', wrap_text=True)

    widths = {
        'A': 9, 'B': 16, 'C': 55, 'D': 10, 'E': 16, 'F': 12, 'G': 24
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 28

    # Keep numbers as numbers and format MRP consistently.
    for cell in ws['F'][1:]:
        if isinstance(cell.value, (int, float)):
            cell.number_format = '0.00'

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def safe_filename(value):
    value = re.sub(r'[^A-Za-z0-9._ -]+', '_', value or 'PO')
    value = re.sub(r'\s+', '_', value).strip('_')
    return value or 'PO'


HTML = '''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>PO PDF → Excel</title>
<style>
*{box-sizing:border-box}body{margin:0;background:#f4f6fa;color:#172033;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif}
.wrap{max-width:1050px;margin:0 auto;padding:48px 20px}.hero{text-align:center;margin-bottom:30px}.hero h1{font-size:48px;margin:0 0 12px;font-weight:800}.hero p{font-size:20px;color:#66738b;margin:0}
.card{background:#fff;border-radius:24px;padding:30px;box-shadow:0 8px 30px rgba(23,32,51,.08);margin-bottom:22px}.drop{border:3px dashed #9aa7bd;border-radius:18px;padding:70px 20px;text-align:center;cursor:pointer}.drop strong{font-size:26px}.drop span{display:block;color:#66738b;font-size:18px;margin-top:12px}input{display:none}
.files{margin:22px 0}.file{padding:14px 4px;border-bottom:1px solid #e7ebf1;display:flex;justify-content:space-between;gap:15px}.btn{width:100%;border:0;background:#172033;color:white;font-size:20px;font-weight:700;padding:18px;border-radius:16px;cursor:pointer}.btn:disabled{opacity:.5}.error{margin-top:18px;color:#e03131;font-size:17px;line-height:1.45}.ok{margin-top:18px;color:#16794b;font-size:17px;line-height:1.45}.info h2{font-size:27px;margin-top:0}.info p{font-size:17px;color:#52617a;line-height:1.7}.pill{background:#f0f3f8;border-radius:7px;padding:2px 7px;font-family:monospace}
</style></head><body><main class="wrap"><div class="hero"><h1>PO PDF → Excel</h1><p>Upload one or more purchase-order PDFs and create a separate Excel file for each PDF.</p></div>
<form class="card" method="post" action="/convert" enctype="multipart/form-data"><label class="drop" for="files"><strong>Tap here to select PO PDF files</strong><span>iPhone · iPad · Android · PC</span></label><input id="files" name="files" type="file" accept="application/pdf,.pdf" multiple><div class="files" id="fileList"></div><button class="btn" type="submit">Convert to Excel</button>
{% if errors %}<div class="error">{% for e in errors %}<div>Error: {{e}}</div>{% endfor %}</div>{% endif %}{% if success %}<div class="ok">{{success}}</div>{% endif %}</form>
<section class="card info"><h2>Excel columns</h2><p><b>SR No · Barcode · Product Description · Qty · PO Number · MRP · Category Name</b></p><p><b>Barcode rule:</b> the T-code immediately after <span class="pill">/</span>, e.g. <span class="pill">T0017521</span>.</p><p><b>MRP:</b> taken from the 3-digit MRP embedded in the PO product description, including values split across PDF lines.</p><p>Each PDF creates its own Excel file named <span class="pill">Vendor_Name_PO_Number.xlsx</span>. When multiple PDFs are selected, the app downloads one ZIP containing the separate Excel files.</p></section></main>
<script>const input=document.getElementById('files'), list=document.getElementById('fileList');input.addEventListener('change',()=>{list.innerHTML='';[...input.files].forEach(f=>{const d=document.createElement('div');d.className='file';d.innerHTML='<span>'+f.name+'</span><span>'+Math.round(f.size/1024/1024*10)/10+' MB</span>';list.appendChild(d)})});</script></body></html>'''


@app.get('/')
def index():
    return render_template_string(HTML, errors=None, success=None)


@app.post('/convert')
def convert():
    files = request.files.getlist('files')
    files = [f for f in files if f and f.filename]
    if not files:
        return render_template_string(HTML, errors=['Please select at least one PDF file.'], success=None), 400

    results = []
    errors = []
    for f in files:
        try:
            data = f.read()
            vendor, po, category, rows = parse_pdf(data)
            xlsx = make_excel(vendor, po, category, rows)
            filename = f'{safe_filename(vendor)}_{safe_filename(po)}.xlsx'
            results.append((filename, xlsx.getvalue()))
        except Exception as exc:
            errors.append(f'{f.filename}: {exc}')

    if not results:
        return render_template_string(HTML, errors=errors, success=None), 422

    if len(results) == 1:
        filename, data = results[0]
        return send_file(io.BytesIO(data), as_attachment=True, download_name=filename,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
        for filename, data in results:
            zf.writestr(filename, data)
    zip_buffer.seek(0)
    return send_file(zip_buffer, as_attachment=True, download_name='PO_Excel_Files.zip', mimetype='application/zip')


@app.get('/health')
def health():
    return {'status': 'ok'}


if __name__ == '__main__':
    port = int(os.environ.get('PORT', '10000'))
    app.run(host='0.0.0.0', port=port)

