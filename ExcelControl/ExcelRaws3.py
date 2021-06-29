import win32com.client
import glob
import os

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

TARGET_FILE = HOME_PATH + '\\Desktop\\TEST.xlsx'

START_LINE = 0
END_LINE = 0

Column = {"COL_A": 1,
          "COL_B": 2,
          "COL_C": 3,
          "COL_D": 4,
          "COL_E": 5,
          "COL_F": 6,
          "COL_G": 7,
          "COL_H": 8,
          "COL_I": 9,
          "COL_J": 10,
          "COL_K": 11,
          "COL_L": 12,
          "COL_M": 13,
          "COL_N": 14,
          "COL_O": 15,
          "COL_P": 16,
          "COL_Q": 17,
          "COL_R": 18,
          "COL_S": 19,
          "COL_T": 20}

ROW_NUM = 2
AP_IP = 120

LOG_FILE = "excel.log"
RES_FILE = "RES.txt"

LOG_DATA = ''

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(TARGET_FILE)
ws = wb.ActiveSheet

# wb2 = excel.Workbooks.Add()
# ws2 = wb2.Worksheets("Sheet1")

if START_LINE == 0:
    startRow = 2
else:
    startRow = START_LINE

if END_LINE == 0:
    totalRows = ws.UsedRange.Rows.Count
else:
    totalRows = END_LINE

firstRow = startRow
lastRow = totalRows

# open('test.txt', 'a', -1, 'utf-8')
f = open(DST_PATH + RES_FILE, "w+", -1, 'utf-8')
attribute = "no;class1;class2;class3;건물;층;설치교실;석면여부\n"
f.write(attribute)

for i in range(firstRow, lastRow+1):
    print(f'{str(i-1)}/{str(lastRow-1)}')
    for k in range(0, int(ws.Cells(i, Column["COL_G"]).value)):

        string = f'{str(int(ws.Cells(i, Column["COL_A"]).value))};' \
                 f'{ws.Cells(i, Column["COL_B"]).value};' \
                 f'{ws.Cells(i, Column["COL_C"]).value};' \
                 f'{ws.Cells(i, Column["COL_D"]).value};' \
                 f'{ws.Cells(i, Column["COL_E"]).value};' \
                 f'{ws.Cells(i, Column["COL_F"]).value};' \
                 f'{ws.Cells(i, Column["COL_H"]).value};' \
                 f'{ws.Cells(i, Column["COL_I"]).value}\n'
        f.write(string)

f.close()
excel.Quit()

