import pandas as pd
from openpyxl import load_workbook

# Define file paths and sheet names
source_file_path = 'path_to_source_file.xlsx'
target_file_path = 'path_to_target_file.xlsx'
source_sheet_name = 'Sheet1'
target_sheet_name = 'Sheet2'

# Define the column to copy and the insertion point in the target sheet
column_to_copy = 'Col'  # The name of the column to copy
insert_after_column = 'Col'  # The column after which the new column will be inserted
insert_before_column = 'Col'  # The column before which the new column will be inserted

# Read the source file and extract the column to copy
source_df = pd.read_excel(source_file_path, sheet_name=source_sheet_name)
column_data = source_df[[column_to_copy]]

# Read the target file
target_df = pd.read_excel(target_file_path, sheet_name=target_sheet_name)

# Check if the column to copy already exists in the target DataFrame
if column_to_copy in target_df.columns:
    # Rename the column to avoid conflict
    new_column_name = f"{column_to_copy}_new"
    column_data.columns = [new_column_name]
else:
    new_column_name = column_to_copy

# Find the position to insert the new column
insert_pos = target_df.columns.get_loc(insert_after_column) + 1

# Insert the column data into the target DataFrame
target_df.insert(insert_pos, new_column_name, column_data)
target_df.drop(columns='Unnamed: 0', inplace=True)

# Save the modified DataFrame back to the Excel file
with pd.ExcelWriter(target_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    target_df.to_excel(writer, sheet_name=target_sheet_name, index=False)

print("Column copied and inserted successfully.")
