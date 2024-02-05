import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font


def create_summary_table(file_path):
    #  -----------Load the workbook and Sheet1 into a DataFrame -----------------
    wb = load_workbook(file_path)
    if 'Working Copy' not in wb.sheetnames:
        first_sheet = wb.worksheets[0]
        first_sheet.title = 'Working Copy'
    ws = wb['Working Copy']
    #  -----------End of Loading the workbook and Sheet1 into a DataFrame --------

    # Define a Font object for bold text
    bold_font = Font(bold=True)

    # Extract the data into a DataFrame, assuming headers are in the second row
    data = ws.iter_rows(min_row=2, values_only=True)
    headers = next(data)  # Get the header row
    df = pd.DataFrame(data, columns=headers)

    # Group by 'Name' and 'MFR Name' and sum the 'Net Bookings'
    summary_df = df.groupby(['Name', 'MFR Name']).agg({'Net Bookings': 'sum'}).reset_index()

    # Check if 'Table' sheet exists, if not, create it
    if 'Table' in wb.sheetnames:
        del wb['Table']
    ws_summary = wb.create_sheet('Table')

    # Write headers in the first row
    ws_summary.append(['Customer Names By Region:', 'Sum of Net Bookings:'])

    # Define the fill color for 'Name' rows
    blue_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')
    currency_format = '"$"#,##0.00'

    # Initialize row_num to 2 to start after the header
    row_num = 2

    for name, group in summary_df.groupby('Name'):
        # Write the 'Name' with total 'Net Bookings', apply blue fill and make it bold
        name_cell = ws_summary.cell(row=row_num, column=1, value=name)
        name_cell.fill = blue_fill
        name_cell.font = bold_font  # Apply bold font to 'Name'
        total_net_bookings_cell = ws_summary.cell(row=row_num, column=2, value=group['Net Bookings'].sum())
        total_net_bookings_cell.number_format = currency_format
        total_net_bookings_cell.font = bold_font  # Apply bold font to the sum value
        row_num += 1

        # Write the 'MFR Names' and their 'Net Bookings' indented under the 'Name'
        for _, row in group.iterrows():
            ws_summary.cell(row=row_num, column=1, value='    ' + row['MFR Name'])  # Add indentation
            net_bookings_cell = ws_summary.cell(row=row_num, column=2, value=row['Net Bookings'])
            net_bookings_cell.number_format = currency_format
            row_num += 1  # Increment the row counter for the next entry

        # Calculate the grand total for bolded 'Name' rows
        grand_total = 0
        for row in ws_summary.iter_rows(min_row=2, max_col=2, max_row=ws_summary.max_row):
            # Check if the name cell is bold, if so, add its corresponding Net Bookings to the grand total
            if row[0].font.bold:
                grand_total += row[1].value

        # Add a grand total row at the end
        grand_total_row = ws_summary.max_row + 1
        ws_summary.cell(row=grand_total_row, column=1, value="Grand Total").font = Font(bold=True)
        grand_total_cell = ws_summary.cell(row=grand_total_row, column=2, value=grand_total)
        grand_total_cell.number_format = currency_format
        grand_total_cell.font = Font(bold=True)  # Make the grand total bold

    # Auto-adjust the column width to fit the content
    for col in ws_summary.columns:
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        adjusted_width = (max_length + 2) * 1.2
        ws_summary.column_dimensions[get_column_letter(col[0].column)].width = adjusted_width

    # Save the workbook
    wb.save(file_path)
