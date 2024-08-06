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

def check_balance_sums(ctb_df):
    """Check if columns C and D sum to zero starting from row 3."""
    # Calculate the sum of column C (Prior Period Balance) and column D (Current Period Balance)
    sum_c = round(ctb_df.iloc[1:, 2].sum(), 2)  # Index 2 corresponds to column C, rounding to 2 decimal places
    sum_d = round(ctb_df.iloc[1:, 3].sum(), 2)  # Index 3 corresponds to column D, rounding to 2 decimal places

    # Print error messages if sums are not zero
    if sum_c != 0:
        print("Prior Period Balance does not sum to 0.")
    if sum_d != 0:
        print("Current Period Balance does not sum to 0.")
    else:
        print("Prior Period Balance and Current Period Balance both sum to 0. ✅")

def process_journal_entries(workbook, sheet_name):
    """Process columns F and G to update values based on the net difference."""
    sheet = workbook[sheet_name]
    last_row = sheet.max_row  # Get the last row based on column A

    # Iterate through each row starting from row 3
    for row in range(3, last_row + 1):
        value_f = sheet.cell(row=row, column=6).value or 0  # Use 0 if cell is None
        value_g = sheet.cell(row=row, column=7).value or 0  # Use 0 if cell is None

        # Calculate the net value
        net_value = value_f - value_g

        if net_value > 0:
            # If net value is positive, update column F and set column G to 0
            sheet.cell(row=row, column=6).value = net_value
            sheet.cell(row=row, column=7).value = 0
        elif net_value < 0:
            # If net value is negative, set column F to 0 and update column G
            sheet.cell(row=row, column=6).value = 0
            sheet.cell(row=row, column=7).value = -net_value

        logging.debug(f"Processed row {row}: F={value_f}, G={value_g}, Net={net_value}")

def check_debit_credit_sums(workbook, sheet_name):
    """Check if the sums of the Debit and Credit columns (F and G) are equal."""
    sheet = workbook[sheet_name]
    last_row = sheet.max_row  # Get the last row based on column A

    # Calculate the sums of columns F and G
    sum_f = sum(sheet.cell(row=row, column=6).value or 0 for row in range(3, last_row + 1))
    sum_g = sum(sheet.cell(row=row, column=7).value or 0 for row in range(3, last_row + 1))

    # Format the sums as strings with accounting format
    sum_f_str = f"${sum_f:,.2f}"
    sum_g_str = f"${sum_g:,.2f}"

    # Check if the sums are equal and print the results
    if round(sum_f, 2) == round(sum_g, 2):
        print(f"The sums of the Debit and Credit columns are equal. Both are {sum_f_str}. ✅")
    else:
        print(f"The sums of the Debit and Credit columns are not equal. The Debit column sums to {sum_f_str} while the Credit column sums to {sum_g_str}. ❌")

def check_journal_entry_balances(workbook, sheet_name, output_directory):
    """Check if each journal entry balances and output the results to an Excel file."""
    sheet = workbook[sheet_name]
    last_row = sheet.max_row

    journal_entries = {}
    for row in range(3, last_row + 1):
        journal_id = sheet.cell(row=row, column=1).value
        date = sheet.cell(row=row, column=3).value
        debit = sheet.cell(row=row, column=6).value or 0
        credit = sheet.cell(row=row, column=7).value or 0

        if journal_id not in journal_entries:
            journal_entries[journal_id] = {"debit": 0, "credit": 0, "dates": set()}

        journal_entries[journal_id]["debit"] += debit
        journal_entries[journal_id]["credit"] += credit
        journal_entries[journal_id]["dates"].add(date)

    je_list_path = os.path.join(output_directory, 'je_list.xlsx')
    je_workbook = openpyxl.Workbook()
    je_sheet = je_workbook.active
    je_sheet.title = "Journal Entries"

    # Add headers
    headers = ["Journal ID", "Net Balance", "Earliest Date", "Latest Date", "Date Difference (Days)"]
    for col_num, header in enumerate(headers, start=1):
        je_sheet.cell(row=1, column=col_num, value=header)

    all_balanced = True
    for row_num, (journal_id, data) in enumerate(journal_entries.items(), start=2):
        net_balance = round(data["debit"] - data["credit"], 2)
        earliest_date = min(data["dates"])
        latest_date = max(data["dates"])
        date_difference = (latest_date - earliest_date).days

        # Check if journal entry is balanced
        if net_balance != 0 or len(data["dates"]) > 1:
            all_balanced = False

        # Add data to the sheet
        je_sheet.cell(row=row_num, column=1, value=journal_id)
        je_sheet.cell(row=row_num, column=2, value=net_balance)
        je_sheet.cell(row=row_num, column=3, value=earliest_date)
        je_sheet.cell(row=row_num, column=4, value=latest_date)
        je_sheet.cell(row=row_num, column=5, value=date_difference)

    # Format column B as Accounting
    accounting_style = NamedStyle(name="accounting_style_je", number_format='_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)')
    if "accounting_style_je" not in je_workbook.named_styles:
        je_workbook.add_named_style(accounting_style)

    for row in range(2, je_sheet.max_row + 1):
        je_sheet.cell(row=row, column=2).style = "accounting_style_je"

    # Adjust column widths
    adjust_column_widths(je_workbook, "Journal Entries")

    # Save the workbook
    je_workbook.save(je_list_path)

    if all_balanced:
        print("All journal entries balance. ✅")
        print("All journal lines for each single journal entry occur on the same date. ✅")
    else:
        print("There are unbalanced journal entries. See 'je_list.xlsx' ❌")
        print("There are journal entries with journal lines on different dates. See 'je_list.xlsx' ❌")

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

    # Load the output workbook to modify data
    workbook = openpyxl.load_workbook(output_file_path)

    # Process columns F and G in 'Journal Entries & Lines'
    process_journal_entries(workbook, "Journal Entries & Lines")

    # Save changes to the workbook
    workbook.save(output_file_path)

    # Validate and highlight the 'Comparative Trial Balances' against the 'Mapping Categories'
    invalid_entries = validate_comparative_trial_balances(df_ctb, df_mapping, output_file_path)

    # Check if columns C and D in 'Comparative Trial Balances' sum to zero
    check_balance_sums(df_ctb)

    # Check if the sums of Debit and Credit columns are equal
    check_debit_credit_sums(workbook, "Journal Entries & Lines")

    # Check if each journal entry balances and write to 'je_list.xlsx'
    check_journal_entry_balances(workbook, "Journal Entries & Lines", current_directory)

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
        print("----COMPARATIVE TRIAL BALANCE TAB----")
        print("All Account Types match one of the options on the Mapping Categories tab. ✅")
        print("All Account Mappings match one of the options on the Mapping Categories tab. ✅")
        print("All Account Mappings have a matching Account Type. ✅")

if __name__ == '__main__':
    main()
