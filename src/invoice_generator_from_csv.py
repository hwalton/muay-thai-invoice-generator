import pandas as pd

# Replace 'yourfile.xlsx' with the path to your XLSX file
file_path = '../data/sessions.xlsx'

# Load the XLSX file as an ExcelFile object
xlsx = pd.ExcelFile(file_path, engine='openpyxl')

# Iterate through each sheet and create a separate CSV file
for sheet_name in xlsx.sheet_names:
    output_file_path = f"../data/{sheet_name.lower().replace(' ', '_')}_csv.csv"  # Generate file name based on sheet name
    data = pd.read_excel(xlsx, sheet_name=sheet_name)
    
    # Write the sheet data to its corresponding CSV file
    data.to_csv(output_file_path, index=False)
    print(f"Data from sheet '{sheet_name}' written to {output_file_path}")

print("All sheets have been processed and saved as CSV files.")


# Load each CSV file into a DataFrame
df_instructions = pd.read_csv('../data/instructions_csv.csv')
df_sessions = pd.read_csv('../data/sessions_csv.csv')
df_instructor_payment_details = pd.read_csv('../data/instructor_payment_details_csv.csv')


# Convert 'Date' to a datetime object
df_sessions['Date'] = pd.to_datetime(df_sessions['Date'])

# Filter out cancelled sessions or those marked with 'script_ignore'
filtered_df_sessions = df_sessions[df_sessions['script_ignore'] != 1]

# Example: Group by 'name_id' and sum the 'Fee' for each instructor
invoice_totals = filtered_df_sessions.groupby('name_id')['Fee'].sum()

# Merge to include instructor payment details
final_invoice_data = invoice_totals.reset_index().merge(df_instructor_payment_details, on='name_id')

# Assuming final_invoice_data is your final DataFrame after all processing
final_invoice_data.to_csv('../data/final_invoices.csv', index=False)

# Load CSV files
df_final_invoices = pd.read_csv('../data/final_invoices.csv')
df_instructor_payment_details = pd.read_csv('../data/instructor_payment_details_csv.csv')
# df_sessions = pd.read_csv('../data/sessions_csv.csv')  # Not used directly in the invoice

# Read the LaTeX template (assuming it's saved as 'template.tex')
with open('../invoices/invoices_template.tex', 'r') as file:
    latex_template = file.read()

# Replace placeholders
for index, row in df_final_invoices.iterrows():
    # Assuming we're creating a single invoice for the first entry in final_invoices.csv
    if index == 0:
        payment_details = df_instructor_payment_details[df_instructor_payment_details['name_id'] == row['name_id']].iloc[0]
        latex_template = latex_template.replace('{{ po_num }}', '12345')  # Replace with actual PO number if available
        latex_template = latex_template.replace('{{ account_name }}', payment_details['account_name'])
        latex_template = latex_template.replace('{{ sort_code }}', payment_details['sort_code'])
        latex_template = latex_template.replace('{{ account_number }}', str(payment_details['account_number']))
        latex_template = latex_template.replace('{{ total_amount }}', str(row['Fee']))

        # Invoice item details - you might want to loop through and add multiple items if needed
        item_description = f"Invoice for {payment_details['first_name']} {payment_details['last_name']}"
        latex_template = latex_template.replace('{{ description }}', item_description)
        latex_template = latex_template.replace('{{ unit_price }}', str(row['Fee']))  # Assuming Fee is the unit price
        latex_template = latex_template.replace('{{ qty }}', '1')  # Quantity is 1 in this example
        latex_template = latex_template.replace('{{ amount }}', str(row['Fee']))  # Total amount for this item
        break  # Remove this if processing multiple invoices

import subprocess

# Write the final LaTeX content to a new file
tex_file_path = '../invoices/filled_invoice.tex'
with open(tex_file_path, 'w') as file:
    file.write(latex_template)

# Compile the LaTeX document using pdflatex
def compile_latex(tex_file):
    process = subprocess.Popen(['pdflatex', tex_file], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    stdout, stderr = process.communicate()

    if process.returncode != 0:
        print(f"Error in LaTeX compilation: {stderr.decode()}")
        return False
    else:
        print(f"LaTeX document compiled successfully. Output in {tex_file.replace('.tex', '.pdf')}")
        return True

# Call the function to compile the LaTeX file
compile_latex(tex_file_path)
