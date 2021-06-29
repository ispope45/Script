import win32com.client
import os

# ### GLOBAL ########
OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

SRC_FILE = 'src.xlsx'
DST_FILE_NAME = 'config'
DST_FILE_EXT = '.txt'

# ### Row Control ######
# 시작과 끝라인 , 0으로 설정하면 맨첫줄(2) 부터 마지막줄까지
START_LINE = 0
END_LINE = 0

# ### COL
# 순번/조직/이름/Hostname/IP/NETMASK/GATEWAY/APC_IP 순으로 배열
NO_COL = 1
ORG_NAME_COL = 2
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
    fileNo = str(ws.Cells(row, NO_COL).Value)
    fileOrg = str(ws.Cells(row, ORG_NAME_COL).Value)
    fileName = str(ws.Cells(row, NAME_COL).Value)

    os.mkdir(DST_PATH + f'{fileNo}_{fileOrg}_{fileName}')
    print(f'{fileNo}_{fileOrg}_{fileName}')

excel.Quit()