# muay-thai-invoice-generator

#python
- `$ venv/bin/activate`
- `$ pip install -r requirements.txt`
- `$ deactivate`

add wicker camp logo as "logo.jpg" to template from google drive/graphics


before requesting PO's:

copy sessions_template.xlsx to sessions.xlsx
copy data from google drive session tracker to sessions.xlsx

run invoice_generator.py script

use the description and payment values to request invoices

once invoice number received:

update PO # received column in sessions.xlsx
save sessions.xlsx

rerun invoice_generator.py script