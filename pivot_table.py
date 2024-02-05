# import openpyxl
# from openpyxl.utils import get_column_letter
# # from openpyxl.workbook.views import SheetView
#
#
# def create_pivot_table(file_path):
#     # Load the workbook
#     wb = openpyxl.load_workbook(file_path)
#
#     # Assuming that the data is in the first sheet
#     data_sheet = wb.worksheets[0]
#
#     # Create a pivot table sheet or select it if it already exists
#     pivot_sheet_title = 'PivotTable'
#     if pivot_sheet_title not in wb.sheetnames:
#         pivot_sheet = wb.create_sheet(pivot_sheet_title)
#     else:
#         pivot_sheet = wb[pivot_sheet_title]
#
#     # Remove any existing pivot tables
#     pivot_sheet._pivots = []
#
#     # Define the range for the Pivot Table data
#     data_range = f'{data_sheet.title}!$A$1:${get_column_letter(data_sheet.max_column)}${data_sheet.max_row}'
#
#     # Create a Pivot Cache Definition
#     pivot_cache_def = openpyxl.pivot.PivotCacheDefinition()
#     pivot_cache_def.cacheSource = pivot_cache_def.create_cacheSource(type="worksheet", worksheet=data_sheet,
#                                                                      ref=data_range)
#     wb.add_pivot_cache(pivot_cache_def)
#
#     # Create the pivot table and add it to the pivot sheet
#     pivot_table = openpyxl.pivot.PivotTable(cache=pivot_cache_def, name="PivotTable1", worksheet=pivot_sheet)
#     pivot_table.add_column(get_column_letter(data_sheet['Name'].column))
#     pivot_table.add_column(get_column_letter(data_sheet['MFR Name'].column))
#     pivot_table.add_data_field(get_column_letter(data_sheet['Net Bookings'].column), name='Sum of Net Bookings',
#                                subtotal='sum')
#
#     pivot_sheet.add_pivot_table(pivot_table)
#
#     # Set the pivot table location
#     pivot_table.location = f'{pivot_sheet.title}!A1'
#
#     # Save the workbook
#     wb.save(file_path)
