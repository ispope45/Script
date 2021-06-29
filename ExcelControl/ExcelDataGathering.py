import win32com.client
import glob
import os

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

TARGET_STRING = "7. 장비현황"
TARGET_FILE = HOME_PATH + '\\Desktop\\dst.xlsx'

START_LINE = 0
END_LINE = 0

COL_A = 1   # 위치
COL_B = 2   # 제조사
COL_C = 3   # 모델명
COL_D = 4
COL_E = 5
COL_F = 6   # 1G
COL_G = 7   # RJ
COL_H = 8   # SFP
COL_I = 9   # 사용망
COL_J = 10  # 비고
COL_K = 11
COL_L = 12

ROW_NUM = 2

files = glob.glob(f'{SRC_PATH}\\*\\*')
print(files)

for file in files:
    if file.find("현장실사점검표") != -1:
        tmp1 = file.split('\\')[5]
        schNum = tmp1.split('_')[0]
        orgName = tmp1.split('_')[1]
        schName = tmp1.split('_')[2]

        table = list

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(file)
        ws = wb.ActiveSheet

        wb2 = excel.Workbooks.Open(TARGET_FILE)
        ws2 = wb2.ActiveSheet

        if START_LINE == 0:
            startRow = 1
        else:
            startRow = START_LINE

        if END_LINE == 0:
            totalRows = ws.UsedRange.Rows.Count
        else:
            totalRows = END_LINE

        firstRow = startRow
        lastRow = totalRows

        for i in range(startRow+1, totalRows):
            colA_data = ws.Cells(i, COL_A).Value

            if colA_data == TARGET_STRING:
                firstRow = i

            if firstRow != startRow and not colA_data:
                lastRow = i

            if firstRow != startRow and lastRow != totalRows:
                break

        print(f'{orgName}_{schName} print line {firstRow} to {lastRow}')

        for row in range(firstRow+2, lastRow):
            ws2.Cells(ROW_NUM, COL_A).Value = schNum
            ws2.Cells(ROW_NUM, COL_B).Value = orgName
            ws2.Cells(ROW_NUM, COL_C).Value = schName
            for col in range(COL_A, COL_J):
                if type(ws.Cells(row, col).Value) == float:
                    data = str(int(ws.Cells(row, col).Value))
                else:
                    data = str(ws.Cells(row, col).Value)
                ws2.Cells(ROW_NUM, col+3).Value = data
                # print(ws2.Cells(ROW_NUM, col+2).Value)
            ROW_NUM += 1

        wb2.Save()
        excel.Quit()

