import openpyxl
from openpyxl.styles import Border, Side

if __name__ == "__main__":
    wb = openpyxl.Workbook()
    ws = wb.active
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    for i in range(1, 10):
        for j in range(1,10):
            ws.cell(column=i, row=j, value=i*j).border = thin_border

    wb.save("border_test.xlsx")