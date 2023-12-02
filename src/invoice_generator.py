import pandas as pd
from jinja2 import Environment, FileSystemLoader
import os
import subprocess

class InvoiceData:
    def __init__(self, name_id, po_num, unit_price, account_name, sort_code, account_number
    ):
        self.name_id = name_id
        self.po_num = po_num
        self.description = 'Instructor for Muay Thai Session(s):'
        self.unit_price = unit_price
        self.qty = 0
        self.amount = self.qty * self.unit_price
        self.total_amount = self.amount
        self.account_name = account_name
        self.sort_code = sort_code
        self.account_number = account_number
        self.BA_dict = {
            'GIAG': 'GIAG',
            'B': 'Beginner',
            'G': 'Graded',
            'Womens GIAG': 'Womens GIAG'
        }

    def update_by_row(self, row):
        self.description += fr'\\{row["Date"]}, {self.BA_dict[row["Beginner/Advanced"]]}'
        assert row['Fee'] == self.unit_price, f'Fee: {row["Fee"]} != unit_price: {self.unit_price} - not constant'
        self.qty += 1
        self.amount = self.qty * self.unit_price
        self.total_amount += self.amount


# Load an Excel file into a pandas DataFrame
df_sessions = pd.read_excel('../data/sessions.xlsx', sheet_name='Sessions')
df_sessions['Date'] = pd.to_datetime(df_sessions['Date']).dt.date
print(f'df: {df_sessions}')

df_bank_details = pd.read_excel('../data/sessions.xlsx', sheet_name='Instructor Payment Details')

# Load the Jinja2 template
env = Environment(loader=FileSystemLoader('../invoices'))
#print(f"os.getcwd: {os.getcwd()}")
template = env.get_template('invoices_template.tex')

unique_name_ids = []


# Filter the DataFrame to include only rows where 'Invoice Sent to SU' is 'NO'
filtered_df = df_sessions[(df_sessions['Invoice Sent to SU'] == 'NO') & (df_sessions['script_ignore'] != 1)]

# Ensure 'Date' is in datetime format



months = [date.month for date in filtered_df['Date'] if pd.notnull(date)]
most_common_month_num = pd.Series(months).mode()[0]

month_dict = {
    1: 'JAN',
    2: 'FEB',
    3: 'MAR',
    4: 'APR',
    5: 'MAY',
    6: 'JUN',
    7: 'JUL',
    8: 'AUG',
    9: 'SEP',
    10: 'OCT',
    11: 'NOV',
    12: 'DEC'
}

most_common_month = month_dict[most_common_month_num]

# Group by 'name_id' and iterate over each group
for _, group in filtered_df.groupby('name_id'):
    first_row = group.iloc[0]  # Select the first row of the group
    data_dict = {
        'name_id': first_row['name_id'],
        'po_num': first_row['PO # Received'],
        'unit_price': first_row['Fee'],
        'most_common_date': most_common_month
    }
    unique_name_ids.append(data_dict)

invoice_list = []

for id in unique_name_ids:
    invoice_data = InvoiceData(
            name_id=id['name_id'],
            po_num = id['po_num'] if not pd.isna(id['po_num']) else fr"{id['name_id']}-{id['most_common_date']}",
            unit_price=id['unit_price'],
            account_name=df_bank_details.loc[df_bank_details['name_id'] == id['name_id'], 'account_name'].values[0],
            sort_code=df_bank_details.loc[df_bank_details['name_id'] == id['name_id'], 'sort_code'].values[0],
            account_number=df_bank_details.loc[df_bank_details['name_id'] == id['name_id'], 'account_number'].values[0]
        )


    invoice_list.append(invoice_data)

for _, row in filtered_df.iterrows():
    for invoice_data in invoice_list:
        if invoice_data.name_id == row['name_id']:
            invoice_data.update_by_row(row)


for invoice_data in invoice_list:
    # Render the template with data from the DataFrame
    rendered_tex = template.render( po_num=invoice_data.po_num,
                                    description=invoice_data.description,
                                    unit_price=invoice_data.unit_price,
                                    qty=invoice_data.qty,
                                    amount=invoice_data.amount,
                                    total_amount=invoice_data.total_amount,
                                    account_name=invoice_data.account_name,
                                    sort_code=invoice_data.sort_code,
                                    account_number=invoice_data.account_number
    )

    # Write the rendered LaTeX to a file
    tex = f'../invoices/invoice-{invoice_data.po_num}'
    tex_tex = f'{tex}.tex'
    with open(tex_tex, 'w') as f:
        f.write(rendered_tex)

    # Compile the LaTeX file to PDF (ensure pdflatex is in your PATH)
    tex_dir = os.path.dirname(os.path.abspath(tex_tex))
    subprocess.run(['pdflatex', os.path.basename(tex_tex)], cwd=tex_dir)

    for ext in ['log', 'aux']:
        os.remove(f'{tex}.{ext}')
