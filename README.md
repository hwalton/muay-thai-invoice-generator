# muay-thai-invoice-generator

# Prerequisites

* python 3.8+
* pip

# Setup
Run:
* `$ python3 -m venv venv`
* `$ source venv/bin/activate`
* `$ pip install -r requirements.txt`

* `$ sudo apt-get install texlive`


* Add wicker camp logo as "logo.jpg" to template from google drive/graphics


Before requesting PO's:
-

* copy sessions_template.xlsx to sessions.xlsx
* Copy data from google drive session tracker to sessions.xlsx
* Run invoice_generator.py script
* Use the description and payment values to request invoices
* Once invoice number received:
* Update PO # received column;
* Rerun invoice_generator.py script to recreate these invoices with PO # included