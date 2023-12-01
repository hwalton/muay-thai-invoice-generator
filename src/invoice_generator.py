import pandas as pd
from jinja2 import Environment, FileSystemLoader
import os
import subprocess

class InvoiceData:
    def __init__(self, po_num, description, unit_price, qty):
        self.po_num = po_num
        self.description = description
        self.unit_price = unit_price
        self.qty = qty
        self.amount = qty * unit_price
        self.total_amount = self.amount


# Load an Excel file into a pandas DataFrame
df = pd.read_excel('../data/sessions.xlsx', sheet_name='Sessions')
print(f'df: {df}')

# Load the Jinja2 template
env = Environment(loader=FileSystemLoader('../invoices'))
#print(f"os.getcwd: {os.getcwd()}")
template = env.get_template('invoices_template.tex')

invoice_list = []

# Create a list of InvoiceData objects
invoice_data = InvoiceData(
    po_num='XXX1',
    description='Desvsnfd',
    unit_price=40,
    qty=1,
)

invoice_list.append(invoice_data)


# Create a list of InvoiceData objects
invoice_data = InvoiceData(
    po_num='XXX2',
    description='Desvsnfd',
    unit_price=40,
    qty=1,
)

invoice_list.append(invoice_data)

for invoice_data in invoice_list:
    # Render the template with data from the DataFrame
    rendered_tex = template.render( po_num=invoice_data.po_num,
                                    description=invoice_data.description,
                                    unit_price=invoice_data.unit_price,
                                    qty=invoice_data.qty,
                                    amount=invoice_data.amount,
                                    total_amount=invoice_data.total_amount
    )

    # Write the rendered LaTeX to a file
    tex = f'../invoices/invoice_{invoice_data.po_num}'
    tex_tex = f'{tex}.tex'
    with open(tex_tex, 'w') as f:
        f.write(rendered_tex)

    # Compile the LaTeX file to PDF (ensure pdflatex is in your PATH)
    tex_dir = os.path.dirname(os.path.abspath(tex_tex))
    subprocess.run(['pdflatex', os.path.basename(tex_tex)], cwd=tex_dir)

    for ext in ['log', 'aux']:
        os.remove(f'{tex}.{ext}')
