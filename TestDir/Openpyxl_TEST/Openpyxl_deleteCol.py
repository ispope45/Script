import openpyxl

if __name__ == "__main__":
    load_wb = openpyxl.load_workbook("C:\\Users\\Jungly\\Downloads\\물품보유현황_20210906_1.xlsx")
    load_ws = load_wb.active

    load_ws.delete_cols(15)

    load_wb.save("C:\\Users\\Jungly\\Downloads\\test2.xlsx")
