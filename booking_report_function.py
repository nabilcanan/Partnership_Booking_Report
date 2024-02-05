from tkinter import filedialog
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00
from tkinter import messagebox
from openpyxl.styles import NamedStyle
from openpyxl.utils import get_column_letter
from table import create_summary_table


# Sheet 1 needs to be 'Working Copy'
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
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    shift_worksheet_down(ws)
    highlight_duplicate_values(ws)
    add_headers_and_formulas(ws)
    format_columns_as_currency_and_percentage(ws)
    format_specific_columns_as_text(ws, ['Qty', 'Ship&Debit'])
    highlight_header_cells(ws)
    freeze_panes(ws)
    format_net_bookings_column(ws)
    add_total_net_bookings(ws)  # Call the new function to add total and highlight
    save_and_notify(file_path, wb)
    create_summary_table(file_path)


def add_total_net_bookings(ws):
    net_bookings_col_index = None

    for col in ws.iter_cols(min_row=2, max_row=2):
        for cell in col:
            if cell.value == 'Net Bookings':
                net_bookings_col_index = cell.column
                break
        if net_bookings_col_index:
            break

    if net_bookings_col_index:
        total_formula = f'=SUM({get_column_letter(net_bookings_col_index)}3:{get_column_letter(net_bookings_col_index)}{ws.max_row})'

        # Check if the style already exists, create a unique name if it does
        style_name = "currency"
        while style_name in ws.parent.named_styles:
            style_name += "_duplicate"

        currency_style = NamedStyle(name=style_name, number_format='"$"#,##0.00')
        ws.parent.add_named_style(currency_style)

        ws.cell(row=1, column=net_bookings_col_index, value=total_formula)
        ws.cell(row=1, column=net_bookings_col_index).style = currency_style

        # Apply the light yellow fill
        light_yellow_fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')
        ws.cell(row=1, column=net_bookings_col_index).fill = light_yellow_fill


def shift_worksheet_down(ws):
    ws.insert_rows(1)


def highlight_duplicate_values(ws):
    seen_values = {}
    lighter_red_fill = PatternFill(start_color='FFDDDD', end_color='FFDDDD', fill_type='solid')

    for row in ws.iter_rows(min_col=1, max_col=1, min_row=2):
        cell = row[0]
        value = cell.value
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


def add_headers_and_formulas(ws):
    ws.cell(row=2, column=13, value='Total Cost')
    ws.cell(row=2, column=14, value='GP%')

    for row in range(3, ws.max_row + 1):
        cost_formula = f'=K{row}-L{row}'
        gp_percentage_formula = f'=(J{row}-M{row})/J{row}'
        ws[f'M{row}'] = cost_formula
        ws[f'N{row}'] = gp_percentage_formula

    ws.insert_cols(23, 4)
    ws.cell(row=2, column=23, value='Notes')
    ws.cell(row=2, column=24, value='Actual Cost')
    ws.cell(row=2, column=25, value='Resale Price')
    ws.cell(row=2, column=26, value='Actual GP%')

    for row in range(3, ws.max_row + 1):
        gp_percentage_formula = f'=(Y{row}-X{row})/Y{row}'
        ws[f'Z{row}'] = gp_percentage_formula
        ws[f'Y{row}'] = ws[f'I{row}'].value


def format_columns_as_currency_and_percentage(ws):
    currency_style = NamedStyle(name='currency', number_format='"$"#,##0.00')
    currency_columns = ['M2:M', 'N2:N', 'Z2:Z', 'Y2:Y', 'K2:K', 'L2:L', 'G2:G', 'H2:H', 'I2:I']

    for cell_range in currency_columns:
        for cell in ws[cell_range + str(ws.max_row)]:
            for col in cell:
                col.style = currency_style

    percentage_style = NamedStyle(name='percentage', number_format=FORMAT_PERCENTAGE_00)
    percentage_columns = ['N2:N', 'Z2:Z']

    for cell_range in percentage_columns:
        for cell in ws[cell_range + str(ws.max_row)]:
            for col in cell:
                col.style = percentage_style


def format_specific_columns_as_text(ws, text_columns):
    for col_name in text_columns:
        col_index = None

        for col in ws.iter_cols(min_row=2, max_row=2):
            for cell in col:
                if cell.value == col_name:
                    col_index = cell.column
                    break
            if col_index:
                break

        if col_index:
            for row in range(3, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_index)
                cell.number_format = '@'  # Format as text


def highlight_header_cells(ws):
    header_cells = ['M2', 'N2', 'W2', 'X2', 'Y2', 'Z2']
    light_yellow_fill = PatternFill(start_color='FFFF99', end_color='FFFF99', fill_type='solid')

    for cell_reference in header_cells:
        ws[cell_reference].fill = light_yellow_fill


def freeze_panes(ws):
    ws.freeze_panes = 'E2'


def format_net_bookings_column(ws):
    net_bookings_col_index = None

    for col in ws.iter_cols(min_row=2, max_row=2):
        for cell in col:
            if cell.value == 'Net Bookings':
                net_bookings_col_index = cell.column
                break
        if net_bookings_col_index:
            break

    if net_bookings_col_index:
        for row in range(3, ws.max_row + 1):
            cell = ws.cell(row=row, column=net_bookings_col_index)
            cell.number_format = '"$"#,##0.00'
    else:
        print("Column 'Net Bookings' not found")


def save_and_notify(file_path, wb):
    wb.save(file_path)
    messagebox.showinfo("Processing Complete", "Excel file processing is complete. You can now save the file.")
