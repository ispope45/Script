import win32com.client
import os

# COMMENT #######
# 포티게이트 URL Filter 설정파일 생성 툴(200818)

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
WEB_ID_COL = 1
URL_ID_COL = 2
URL_CONTENTS_COL = 3

URL_TYPE_COL = 4
URL_ACTION_COL = 5

COMMENT_COL = 6

# 추가 멘트
# 설정값 맨 앞줄에 추가
IS_FIRST_COMMENT = True
FIRST_COMMENT = 'config webfilter urlfilter\n' \
                'config edit 2\n' \
                'config entries\n'

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

    webFilterId = str(ws.Cells(row, WEB_ID_COL).Value)
    urlFilterId = str(ws.Cells(row, URL_ID_COL).Value)
    urlFilterContents = str(ws.Cells(row, URL_CONTENTS_COL).Value)
    urlFilterType = str(ws.Cells(row, URL_TYPE_COL).Value)
    urlFilterAction = str(ws.Cells(row, URL_ACTION_COL).Value)

    if row == START_LINE+2:
        script = (
            f'{FIRST_COMMENT}'
        )
    script = (
        f'edit {urlFilterId}\n'
        f'set url "{urlFilterContents}"\n'
        f'set type {urlFilterType}\n'
        f'set action {urlFilterAction}\n'
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