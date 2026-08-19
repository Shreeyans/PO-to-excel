# PO PDF → Excel Web App

Cross-platform web app for iPhone, iPad, Android and PC.

Workflow:
1. Select one or more PO PDFs.
2. Each PDF is processed separately.
3. Extract:
   - SR No
   - Barcode = T-code immediately after `/`
   - Product Description
   - Qty
   - PO Number
   - MRP
   - Category Name
4. Download one Excel per PDF.
5. Filename: Vendor Name_PO Number.xlsx

## Run

Requires Python 3.10+.

    pip install -r requirements.txt
    python app.py

Open the displayed address from any device on the same network, or deploy the app to a web host.

## Important

The browser app sends the selected PDFs to the Python server for processing. For production use, add HTTPS, authentication, file-size limits and automatic deletion of uploaded PDFs.
