import openpyxl

def write_to_excel(excel_name, tab_name, row_num, col_num):
    wb = openpyxl.load_workbook(excel_name)
    sheet = wb[tab_name]
    cell = sheet.cell(row=row_num, column=col_num)
    cell.value = "Hello World"
    wb.save(excel_name)
    
    
    
