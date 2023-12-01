import pandas as pd

# Load an Excel file into a pandas DataFrame
df = pd.read_excel('../data/sessions.xlsx', sheet_name='Sessions')

print(f"df: {df}")