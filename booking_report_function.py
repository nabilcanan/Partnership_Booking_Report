from tkinter import filedialog
import openpyxl
from openpyxl.styles import PatternFill, NamedStyle
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
from tkinter import messagebox


def on_file_selected():
    file_path = select_excel_file()
    if file_path:
        process_excel_file(file_path)


def select_excel_file():
    # Show an "Open" dialog box and return the path to the selected file
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:  # Check if a file was selected
        print(f"File selected: {file_path}")
        return file_path
    else:
        print("No file was selected.")
        return None


def process_excel_file(file_path):
    # Load the workbook and select the active worksheet
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    # Dictionary to track duplicates in column A
    seen_values = {}

    # Red fill for highlighting duplicates
    lighter_red_fill = PatternFill(start_color='FFDDDD', end_color='FFDDDD', fill_type='solid')

    # Iterate through column A to find duplicates and highlight them
    for row in ws.iter_rows(min_col=1, max_col=1, min_row=2):  # Assuming row 1 is headers
        cell = row[0]
        value = cell.value
        # Strip leading zeros
        value = value.lstrip('0') if isinstance(value, str) else value
        cell.value = value
        if value in seen_values:
            seen_values[value].append(cell.coordinate)
        else:
            seen_values[value] = [cell.coordinate]

    for duplicates in seen_values.values():
        if len(duplicates) > 1:
            for coordinate in duplicates:
                ws[coordinate].fill = lighter_red_fill

    # 2. Shift existing data in columns M and N to O and P, and create new blank columns M and N
    ws.insert_cols(13, 2)  # Inserts two columns starting at column 13 (which is M), don't need to shift
    # This line ^^ makes this it, so we create new columns where we are adding the new columns

    # Add headers for columns M and N
    ws.cell(row=1, column=13, value='Total Cost')
    ws.cell(row=1, column=14, value='GP%')

    # Add formulas to columns M and N using the 'formula' attribute
    for row in range(2, ws.max_row + 1):
        cost_formula = f'=K{row}-L{row}'
        gp_percentage_formula = f'=(J{row}-M{row})/J{row}'
        ws[f'M{row}'] = cost_formula
        ws[f'N{row}'] = gp_percentage_formula

    # Add columns W, X, Y, and Z with headers and formulas
    ws.insert_cols(23, 4)  # Inserts four columns starting at column 23 (W)

    # Add headers for columns W, X, Y, and Z
    ws.cell(row=1, column=23, value='Notes')
    ws.cell(row=1, column=24, value='Actual Cost')
    ws.cell(row=1, column=25, value='Resale Price')
    ws.cell(row=1, column=26, value='Actual GP%')

    # Add formulas and data to columns W, X, Y, and Z using the 'formula' attribute
    for row in range(2, ws.max_row + 1):
        # Formulas for columns Z (Actual GP%) and Y (Resale Price)
        gp_percentage_formula = f'=(Y{row}-X{row})/Y{row}'
        ws[f'Z{row}'] = gp_percentage_formula

        # Copy data from column I to column Y
        ws[f'Y{row}'] = ws[f'I{row}'].value

        # Formulas for columns W (Notes) and X (Actual Cost)
        # These columns will remain blank as per your request
    # Format column N (GP%) and column Z (Actual GP%) as percentages
    percentage_style = NamedStyle(name='percentage', number_format=FORMAT_PERCENTAGE_00)
    for cell_range in ['N2:N', 'Z2:Z']:  # Update the column ranges as needed
        for cell in ws[cell_range + str(ws.max_row)]:
            for col in cell:
                col.style = percentage_style

    # Merge in here freeze panes and function for net bookings as dollar sign and .00 decimal point

    # Highlight headers in columns M, N, W, X, Y, and Z with light yellow fill
    header_cells = ['M1', 'N1', 'W1', 'X1', 'Y1', 'Z1']
    light_yellow_fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
    for cell_reference in header_cells:
        ws[cell_reference].fill = light_yellow_fill

        # Freeze panes at E2
        ws.freeze_panes = 'E2'

        # Assuming 'ws' is your worksheet object
        net_bookings_col_index = None

        # Locate the column with the header "Net Bookings"
        for col in ws.iter_cols(min_row=1, max_row=1):  # Search in the first row
            for cell in col:
                if cell.value == 'Net Bookings':
                    net_bookings_col_index = cell.column
                    break
            if net_bookings_col_index:
                break

        # Check if "Net Bookings" column was found
        if net_bookings_col_index:
            # Format the 'Net Bookings' column with dollar sign and two decimal places
            for row in range(2, ws.max_row + 1):  # Start from row 2 to skip the header
                cell = ws.cell(row=row, column=net_bookings_col_index)
                cell.number_format = '"$"#,##0.00'
        else:
            print("Column 'Net Bookings' not found")

    # Save the modified workbook
    wb.save(file_path)

    # Notify the user and provide an option to save the file
    messagebox.showinfo("Processing Complete", "Excel file processing is complete. You can now save the file.")
