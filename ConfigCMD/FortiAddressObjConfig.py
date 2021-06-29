import win32com.client
import os

# COMMENT #######
# 포티게이트 URL Filter 설정파일 생성 툴(200818)

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

NAME_COL = 1
IP_COL = 2
INTERFACE_COL = 3


# 추가 멘트
# 설정값 맨 앞줄에 추가
IS_FIRST_COMMENT = False
FIRST_COMMENT = 'config firewall address\n'

# 설정값 맨 마지막줄에 추가
IS_EXTRA_COMMENT = False
EXTRA_COMMENT = '\n\n'

# ###

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(HOME_PATH + SRC_FILE)
ws = wb.ActiveSheet

content = ''

if START_LINE == 0:
    START_LINE = 1

if END_LINE == 0:
    END_LINE = ws.UsedRange.Rows.Count

for row in range(START_LINE+1, END_LINE+1):
    fileName = DST_FILE_NAME + DST_FILE_EXT

    addressName = str(ws.Cells(row, NAME_COL).Value)
    addressIp = str(ws.Cells(row, IP_COL).Value)
    addressInt = str(ws.Cells(row, INTERFACE_COL).Value)

    if row == START_LINE+2:
        script = (
            f'{FIRST_COMMENT}'
        )
    script = (
        f'edit "{addressIp}_{addressName}"\n'
        f'set associated-interface "{addressInt}"\n'
        f'set subnet {addressIp} 255.255.255.255\n'
        f'next\n'
    )

    if IS_EXTRA_COMMENT:
        script = script + EXTRA_COMMENT

    content = content + script
    content = content + '\n'

f = open(DST_PATH + fileName, "w+")
f.write(content)
f.close()

excel.Quit()
