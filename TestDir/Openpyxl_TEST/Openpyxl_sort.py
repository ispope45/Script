import openpyxl
import pandas as pd

if __name__ == "__main__":
    # load_wb = openpyxl.load_workbook("C:\\Users\\Jungly\\Downloads\\물품보유현황_20210906_1.xlsx")
    # load_ws = load_wb.active
    #
    # startRow = 2
    # totalRows = load_ws.max_row + 1
    #
    # load_ws.auto_filter.ref = f'A1:O{str(totalRows)}'
    # load_ws.auto_filter.add_sort_condition(f'B2:B{str(totalRows)}')
    #
    #
    # load_wb.save("C:\\Users\\Jungly\\Downloads\\물품보유현황_20210906_2.xlsx")

    test = pd.read_excel("C:\\Users\\Jungly\\Downloads\\물품보유현황_20210906_1.xlsx")
    print(test)
    test = test.sort_values(by=['운용부서', '물품분류명'], ascending=[True, True])
    print(test)
    test.to_excel("C:\\Users\\Jungly\\Downloads\\물품보유현황_20210906_3.xlsx")