import openpyxl

if __name__ == "__main__":
    load_wb = openpyxl.load_workbook("C:\\Users\\Jungly\\Downloads\\test.xlsx")
    load_ws = load_wb.active

    startRow = 2
    totalRows = load_ws.max_row + 1

    load_ws.auto_filter.ref = f'A1:Q{str(totalRows)}'
    load_ws.auto_filter.add_sort_condition(f'B2:B{str(totalRows)}')

    load_wb.save("C:\\Users\\Jungly\\Downloads\\test1.xlsx")