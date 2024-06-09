import pandas as pd
from jinja2 import Environment, FileSystemLoader
import os
import subprocess
from datetime import datetime
import numpy as np

class InvoiceData:
    def __init__(self, name_id, po_num, unit_price, account_name, sort_code, account_number, first_name, last_name
    ):
        self.name_id = name_id
        self.po_num = po_num
        self.description = f'Instructor {first_name} {last_name} for Muay Thai Session(s):'
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
            'M': 'Mixed',
            'GS': 'Graded Sparring',
            'GT': 'Graded Technique',
            'Womens GIAG': 'Womens GIAG',
            'W': 'Womens'
        }
        self.first_name = first_name
        self.last_name = last_name

    def update_by_row(self, row):
        self.description += fr'\\{row["Date"].strftime("%d-%m-%Y")} {row["Beginner/Advanced"]};'
        assert row['Fee'] == self.unit_price, f'Fee: {row["Fee"]} != unit_price: {self.unit_price} - not constant'
        self.qty += 1
        self.amount = self.qty * self.unit_price
        self.total_amount = self.amount


# Load an Excel file into a pandas DataFrame
data_types = {'Week': 'int',
              'Fee': 'float',
              'PO # Received': 'Int64',
              'account_number': 'str',
              'sort_code': 'str'}

df_sessions = pd.read_excel('../data/sessions.xlsx', sheet_name='Sessions', dtype=data_types)
df_sessions['Date'] = pd.to_datetime(df_sessions['Date']).dt.date
df_sessions = df_sessions[(df_sessions['script_ignore'] == 0)]
print(f'df: {df_sessions}')

df_bank_details = pd.read_excel('../data/sessions.xlsx', sheet_name='Instructor Payment Details', dtype=data_types)

# Load the Jinja2 template
env = Environment(loader=FileSystemLoader('../invoices'))
#print(f"os.getcwd: {os.getcwd()}")
template = env.get_template('invoices_template.tex')

unique_name_ids = []

# Asserts for the DataFrame to ensure sessions spreadsheet is correctly formatted
found_no = False
valid_days = {"Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"}
valid_Beginner_Advanced = {"GIAG", "B", "G", "M", "GS", "GT", "Womens GIAG", "W"}
valid_Location = {"Pearson (Hallam)", "Wicker Camp", "Goodwin Matrix Studio"}
valid_name_ids = set(df_bank_details['name_id'])
valid_first_names = set(df_bank_details['first_name'])
valid_last_names = set(df_bank_details['last_name'])
valid_PO_Requested = {"YES", "NO"}
valid_Invoice_Sent_to_SU = {"YES", "NO"}
valid_Payment_Confirmed_Completed = {"YES", "NO"}


for i in range(len(df_sessions)):
    row = df_sessions.iloc[i]

    assert df_sessions.shape[1] == 16, f"The number of columns in the DataFrame is not 16, it's {df_sessions.shape[1]}"

    assert row['script_ignore'] in [0,1], f"Invalid value in 'script_ignore' column at line {i}: {row['script_ignore']}"

    try:
        date_str = str(row['Date'])

        parsed_date = datetime.strptime(date_str, '%Y-%m-%d')
    except ValueError:
        raise AssertionError(
            f"Date format is incorrect at line {i}: {row['Date']} (expected format: 'yyyy-mm-dd')")


    assert row['Day'] in valid_days, f"Invalid day of the week at line {i}: {row['Day']}"

    week = row['Week']
    assert np.issubdtype(type(week), np.integer) and -100 <= week <= 100, f"'Week' value is out of range or not an integer at line {i}: {week}"

    assert row['Beginner/Advanced'] in valid_Beginner_Advanced, f"Invalid value in 'Beginner/Advanced' column at line {i}: {row['Beginner/Advanced']}"

    assert row['Location'] in valid_Location, f"Invalid value in 'Location' column at line {i}: {row['Location']}"

    assert row['name_id'] in valid_name_ids, f"'name_id' value not found in df_bank_details at line {i}: {row['name_id']}"

    assert row['first_name'] in valid_first_names, f"'first_name' value not found in df_bank_details at line {i}: {row['first_name']}"

    assert row['last_name'] in valid_last_names, f"'last_name' value not found in df_bank_details at line {i}: {row['last_name']}"

    fee = row['Fee']
    assert np.issubdtype(type(fee), np.floating) and 0 <= fee <= 100, f"'Fee' value is out of range or not a float at line {i}: {fee}"

    assert row['PO Requested'] in valid_PO_Requested, f"Invalid value in 'PO Requested' column at line {i}: {row['PO Requested']}"

    PO_number = row['PO # Received']
    assert (pd.isna(PO_number) or (np.issubdtype(type(PO_number), np.integer) and 0 <= PO_number <= 100000)), f"'PO # Received' value is out of range, not an integer, or not empty at line {i}: {PO_number}"

    assert row['Invoice Sent to SU'] in valid_Invoice_Sent_to_SU, f"Invalid value in 'Invoice Sent to SU' column at line {i}: {row['Invoice Sent to SU']}"

    assert row['Payment Confirmed Complete'] in valid_Payment_Confirmed_Completed, f"Invalid value in 'Payment Confirmed Complete' column at line {i}: {row['Payment Confirmed Complete']}"

    if row['Invoice Sent to SU'] == 'NO':
        found_no = True

assert found_no, "No new invoices to create (No line contains 'NO' in the 'Invoice Sent to SU' column.)"


filtered_df = df_sessions[(df_sessions['Invoice Sent to SU'] == 'NO') & (df_sessions['script_ignore'] == 0)]

# Ensure 'Date' is in datetime format
months = [date.month for date in filtered_df['Date'] if pd.notnull(date)]
years = [date.year for date in filtered_df['Date'] if pd.notnull(date)]
most_common_month_num = pd.Series(months).mode()[0]
most_common_year = str(pd.Series(years).mode()[0])

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

    # Extract the "PO # Received" column, ignoring NaN for equality check
    po_received = group['PO # Received'].dropna().unique()

    # Check if all values are NaN in the original column
    all_na = group['PO # Received'].isna().all()
    any_na = group['PO # Received'].isna().any()

    # Assert that either there is only one unique value (ignoring NaNs) or all values are NaN
    assert len(po_received) <= 1 and all_na == any_na, "Not all values in 'PO # Received' are the same or NaN for a group"

    first_row = group.iloc[0]  # Select the first row of the group

    data_dict = {
        'name_id': first_row['name_id'],
        'po_num': first_row['PO # Received'],
        'unit_price': first_row['Fee'],
        'most_common_date': most_common_year + '-' + most_common_month
    }
    unique_name_ids.append(data_dict)

invoice_list = []

for id in unique_name_ids:
    invoice_data = InvoiceData(
            name_id=id['name_id'],
            po_num = int(id['po_num']) if not pd.isna(id['po_num']) else fr"{id['most_common_date']}-{id['name_id']}",
            unit_price=id['unit_price'],
            account_name=df_bank_details.loc[df_bank_details['name_id'] == id['name_id'], 'account_name'].values[0],
            sort_code=df_bank_details.loc[df_bank_details['name_id'] == id['name_id'], 'sort_code'].values[0],
            account_number=df_bank_details.loc[df_bank_details['name_id'] == id['name_id'], 'account_number'].values[0],
            first_name=df_bank_details.loc[df_bank_details['name_id'] == id['name_id'], 'first_name'].values[0],
            last_name=df_bank_details.loc[df_bank_details['name_id'] == id['name_id'], 'last_name'].values[0]
        )


    invoice_list.append(invoice_data)

for _, row in filtered_df.iterrows():
    for invoice_data in invoice_list:
        if invoice_data.name_id == row['name_id']:
            invoice_data.update_by_row(row)


for invoice_data in invoice_list:
    # Render the template with data from the DataFrame
    rendered_tex = template.render( po_num=invoice_data.po_num,
                                    description=str(invoice_data.description),
                                    unit_price="{:.2f}".format(invoice_data.unit_price),
                                    qty=int(invoice_data.qty),
                                    amount="{:.2f}".format(invoice_data.amount),
                                    total_amount="{:.2f}".format(invoice_data.total_amount),
                                    account_name=str(invoice_data.account_name),
                                    sort_code=str(invoice_data.sort_code),
                                    account_number=str(invoice_data.account_number),
                                    first_name=str(invoice_data.first_name),
                                    last_name=str(invoice_data.last_name)
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
