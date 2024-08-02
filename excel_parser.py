import pandas as pd
import openpyxl
import os
from openpyxl.styles import PatternFill, Font, Alignment, NamedStyle
import logging

# Set up logging for debugging
logging.basicConfig(level=logging.INFO)

def find_last_row(input_file_path, sheet_name):
    """Find the last consecutive non-empty row in column A with data, starting from the top."""
    # Load the workbook and select the specified sheet
    workbook = openpyxl.load_workbook(input_file_path, data_only=True)
    sheet = workbook[sheet_name]
    
    # Iterate from the top down to find the last consecutive non-empty cell in column A
    for row in range(1, sheet.max_row + 1):
        if sheet.cell(row=row, column=1).value is None:
            return row - 1
    return sheet.max_row  # If no empty row is found, return max_row

def read_excel(input_file_path, sheet_name, last_row):
    """Read data from specific columns up to the last non-empty row in column A."""
    # Read columns A to G from the specified sheet up to last_row
    df = pd.read_excel(
        input_file_path, 
        sheet_name=sheet_name, 
        usecols="A:G", 
        nrows=last_row,  # Use last_row directly for nrows
        engine='openpyxl'
    )
    return df

def read_full_sheet(input_file_path, sheet_name):
    """Read all data from a specified sheet."""
    df = pd.read_excel(
        input_file_path,
        sheet_name=sheet_name,
        engine='openpyxl'
    )
    return df

def write_excel(df1, df2, df3, output_file_path):
    """Write DataFrames to a new Excel file with specific sheets and format headers."""
    # Create a new workbook and add sheets
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        # Create a blank 'Instructions' sheet
        pd.DataFrame().to_excel(writer, sheet_name='Instructions', index=False)

        # Create a blank 'Data Validation Tests' sheet
        pd.DataFrame().to_excel(writer, sheet_name='Data Validation Tests', index=False)

        # Write the first DataFrame to the 'Comparative Trial Balances' sheet
        df1.to_excel(writer, sheet_name='Comparative Trial Balances', index=False)

        # Write the second DataFrame to the 'Journal Entries & Lines' sheet
        df2.to_excel(writer, sheet_name='Journal Entries & Lines', index=False)

        # Write the third DataFrame to the 'Mapping Categories' sheet
        df3.to_excel(writer, sheet_name='Mapping Categories', index=False)

    # Load the workbook again to apply styles
    workbook = openpyxl.load_workbook(output_file_path)

    # Apply styles to the 'Comparative Trial Balances' header and second row
    format_header_and_second_row(workbook, "Comparative Trial Balances")

    # Apply styles to the 'Journal Entries & Lines' header and specific second row columns
    format_header_and_second_row_jel(workbook, "Journal Entries & Lines")

    # Apply accounting format to columns C and D of 'Comparative Trial Balances'
    apply_accounting_format(workbook, "Comparative Trial Balances", [3, 4])

    # Apply accounting and date formats to 'Journal Entries & Lines'
    apply_accounting_and_date_format_jel(workbook, "Journal Entries & Lines")

    # Adjust column widths for readability
    adjust_column_widths(workbook, "Comparative Trial Balances")
    adjust_column_widths(workbook, "Journal Entries & Lines")

    # Save changes to the workbook
    workbook.save(output_file_path)

def format_header_and_second_row(workbook, sheet_name):
    """Apply formatting to the header and second row in the specified sheet."""
    sheet = workbook[sheet_name]
    
    # Define styles
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")  # Dark blue fill
    second_row_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")  # Light blue fill
    header_font = Font(color="FFFFFF", bold=True)  # White bold font
    center_alignment = Alignment(horizontal="center")  # Center alignment

    # Apply styles to the first row, columns A:G
    for col in range(1, 8):  # Columns A:G are 1:7 in 1-indexed systems
        cell = sheet.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment

    # Apply styles to the second row, columns A:G
    for col in range(1, 8):
        cell = sheet.cell(row=2, column=col)
        cell.fill = second_row_fill
        cell.font = header_font
        cell.alignment = center_alignment

def format_header_and_second_row_jel(workbook, sheet_name):
    """Apply specific formatting to the header and second row in the 'Journal Entries & Lines' sheet."""
    sheet = workbook[sheet_name]
    
    # Define styles
    header_fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")  # Dark blue fill
    light_blue_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")  # Light blue fill
    grey_fill = PatternFill(start_color="999999", end_color="999999", fill_type="solid")  # Grey fill
    header_font = Font(color="FFFFFF", bold=True)  # White bold font
    center_alignment = Alignment(horizontal="center")  # Center alignment

    # Apply styles to the first row, columns A:G
    for col in range(1, 8):  # Columns A:G are 1:7 in 1-indexed systems
        cell = sheet.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_alignment

    # Apply light blue styles to the second row, columns A, C, D, F, G
    for col in [1, 3, 4, 6, 7]:
        cell = sheet.cell(row=2, column=col)
        cell.fill = light_blue_fill
        cell.font = header_font
        cell.alignment = center_alignment

    # Apply grey styles to the second row, columns B, E
    for col in [2, 5]:
        cell = sheet.cell(row=2, column=col)
        cell.fill = grey_fill
        cell.font = header_font
        cell.alignment = center_alignment

def apply_accounting_format(workbook, sheet_name, columns):
    """Apply accounting number format to specified columns starting from row 3."""
    sheet = workbook[sheet_name]
    accounting_style = NamedStyle(name="accounting_style", number_format='_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)')

    # Check if the style already exists to avoid duplication error
    if "accounting_style" not in workbook.named_styles:
        workbook.add_named_style(accounting_style)
    else:
        logging.info("Accounting style already exists.")

    # Apply accounting style to specified columns starting from row 3
    for col in columns:
        for row in range(3, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col)
            cell.style = "accounting_style"
            logging.debug(f"Applied accounting style to {sheet.title}!{cell.coordinate}")

def apply_accounting_and_date_format_jel(workbook, sheet_name):
    """Apply accounting number format to columns F and G, and date format to column C starting from row 3."""
    sheet = workbook[sheet_name]
    accounting_style = NamedStyle(name="accounting_style", number_format='_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)')
    date_style = NamedStyle(name="date_style", number_format='m/d/yy')

    # Add styles to workbook if they don't already exist
    if "accounting_style" not in workbook.named_styles:
        workbook.add_named_style(accounting_style)
    else:
        logging.info("Accounting style already exists in 'Journal Entries & Lines'.")

    if "date_style" not in workbook.named_styles:
        workbook.add_named_style(date_style)
    else:
        logging.info("Date style already exists.")

    # Apply accounting style to columns F and G starting from row 3
    for col in [6, 7]:
        for row in range(3, sheet.max_row + 1):
            cell = sheet.cell(row=row, column=col)
            cell.style = "accounting_style"
            logging.debug(f"Applied accounting style to {sheet.title}!{cell.coordinate}")

    # Apply date style to column C starting from row 3
    for row in range(3, sheet.max_row + 1):
        cell = sheet.cell(row=row, column=3)
        cell.style = "date_style"
        logging.debug(f"Applied date style to {sheet.title}!{cell.coordinate}")

def adjust_column_widths(workbook, sheet_name):
    """Adjust column widths to fit the content in the specified sheet."""
    sheet = workbook[sheet_name]

    # Calculate the maximum width needed for each column based on its content
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column letter
        for cell in col:
            try:
                # Update max_length if current cell is longer
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max_length + 2  # Add some padding to avoid cutting off text
        sheet.column_dimensions[column].width = adjusted_width
        logging.debug(f"Adjusted width of column {column} to {adjusted_width}")

def validate_comparative_trial_balances(ctb_df, mapping_df, output_file_path):
    """Validate and highlight that each value in column E of 'Comparative Trial Balances' is in column A of 'Mapping Categories'."""
    # Get the set of valid mapping categories, ensuring consistent formatting
    valid_categories = set(mapping_df.iloc[0:, 0].dropna().astype(str).str.strip())  # Convert to string and strip whitespace
    
    # Open the output file to apply highlights
    workbook = openpyxl.load_workbook(output_file_path)
    sheet = workbook["Comparative Trial Balances"]
    
    # Define the fill for highlighting
    highlight_fill_e = PatternFill(start_color="F7B4AE", end_color="F7B4AE", fill_type="solid")  # Pink fill for column E
    highlight_fill_f = PatternFill(start_color="F7B4AE", end_color="F7B4AE", fill_type="solid")  # Pink fill for column F

    # Check each value in column E starting from row 3
    invalid_entries = []
    for index, (category, value) in enumerate(zip(ctb_df.iloc[2:, 4].dropna().astype(str).str.strip(),
                                                  ctb_df.iloc[2:, 5].dropna().astype(str).str.strip()), start=3):
        if category not in valid_categories:
            # If column E is invalid, highlight column E and F
            invalid_entries.append((index + 1, category))
            sheet.cell(row=index + 1, column=5).fill = highlight_fill_e
            sheet.cell(row=index + 1, column=6).fill = highlight_fill_f
        else:
            # If column E is valid, validate column F based on the mapping
            column_mapping = {
                "Assets": 1,        # Validate against column B (0-indexed)
                "Liabilities": 2,   # Validate against column C
                "Equity": 3,        # Validate against column D
                "Income": 4,        # Validate against column E
                "Expenses": 5       # Validate against column F
            }
            mapping_column = column_mapping.get(category)
            if mapping_column is not None:
                valid_values = set(mapping_df.iloc[0:, mapping_column].dropna().astype(str).str.strip())
                if value not in valid_values:
                    invalid_entries.append((index + 1, category, value))
                    sheet.cell(row=index + 1, column=6).fill = highlight_fill_f

    # Save changes to the workbook
    workbook.save(output_file_path)
    
    return invalid_entries

def main():
    # Define file names
    input_file_name = 'source_file.xlsx'  # Replace with your actual file name
    output_file_name = 'output_file.xlsx' # Name for the output file

    # Get the directory of the current script
    current_directory = os.path.dirname(os.path.abspath(__file__))

    # Construct full file paths
    input_file_path = os.path.join(current_directory, input_file_name)
    output_file_path = os.path.join(current_directory, output_file_name)

    # Find the last row in column A with data for the 'Comparative Trial Balances' sheet
    last_row_ctb = find_last_row(input_file_path, "Comparative Trial Balances")

    # Read the Excel file up to the last row for the 'Comparative Trial Balances' sheet
    df_ctb = read_excel(input_file_path, "Comparative Trial Balances", last_row_ctb)

    # Find the last row in column A with data for the 'Journal Entries & Lines' sheet
    last_row_jel = find_last_row(input_file_path, "Journal Entries & Lines")

    # Read the Excel file up to the last row for the 'Journal Entries & Lines' sheet
    df_jel = read_excel(input_file_path, "Journal Entries & Lines", last_row_jel)

    # Read the entire 'Mapping Categories' sheet
    df_mapping = read_full_sheet(input_file_path, "Mapping Categories")

    # Write the data to a new Excel file with additional sheets
    write_excel(df_ctb, df_jel, df_mapping, output_file_path)

    # Validate and highlight the 'Comparative Trial Balances' against the 'Mapping Categories'
    invalid_entries = validate_comparative_trial_balances(df_ctb, df_mapping, output_file_path)

    # Print validation results
    if invalid_entries:
        print("Invalid entries found in 'Comparative Trial Balances':")
        for entry in invalid_entries:
            if len(entry) == 2:
                row, category = entry
                print(f"Row {row}: Category '{category}' not found in 'Mapping Categories'")
            else:
                row, category, value = entry
                print(f"Row {row}: Value '{value}' not valid for category '{category}' in 'Mapping Categories'")
    else:
        print("All entries in 'Comparative Trial Balances' are valid.")

if __name__ == '__main__':
    main()
