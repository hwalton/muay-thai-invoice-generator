# muay-thai-invoice-generator

activate venv with:
$source venv/bin/activate

before requesting PO's:

copy data from google drive session tracker to sessions.xlsx

run invoice_generator.py script

use the description and payment values to request invoices

once invoice number received:

update PO # received column;
rerun invoice_generator.py script