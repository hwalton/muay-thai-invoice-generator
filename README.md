# muay-thai-invoice-generator

#python
- `$ venv/bin/activate`
- `$ pip install -r requirements.txt`
- `$ deactivate`


before requesting PO's:

copy sessions_template.xlsx to sessions.xlsx
copy data from google drive session tracker to sessions.xlsx

run invoice_generator.py script

use the description and payment values to request invoices

once invoice number received:

update PO # received column;
rerun invoice_generator.py script