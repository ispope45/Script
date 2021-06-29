import win32com.client
import os

# ### GLOBAL ########
OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH + '\\Desktop\\'

SRC_PATH = HOME_PATH + 'src\\'
DST_PATH = HOME_PATH + 'dst\\'

SRC_FILE = 'src.xlsx'
DST_FILE_NAME = 'config'
DST_FILE_EXT = '.txt'

# ### Row Control ######
# 시작과 끝라인 , 0으로 설정하면 맨첫줄(2) 부터 마지막줄까지
START_LINE = 0
END_LINE = 0

# ### COL
FROM = 1
TO = 2
NAME_COL = 3

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(HOME_PATH + SRC_FILE)
ws = wb.ActiveSheet

if START_LINE == 0:
    START_LINE = 1

if END_LINE == 0:
    END_LINE = ws.UsedRange.Rows.Count

for row in range(START_LINE+1, END_LINE+1):
    dirFrom = str(ws.Cells(row, FROM).Value)
    dirTo = str(ws.Cells(row, TO).Value)

    # os.mkdir(DST_PATH + f'{fileNo}_{fileOrg}_{fileName}')
    print(f'{dirFrom} :: {dirTo}')
    if not (os.path.isdir(dirTo)):
        os.rename(dirFrom, dirTo)

excel.Quit()
