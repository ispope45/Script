import win32com.client
import glob
import os
# COMMENT #######
# src 디렉터리에 있는 엑셀파일(xlsx)들을 대상으로 특정 시트이름(Sheet1 등)을 가진 시트를 삭제후
# dst 디렉터리에 저장

# HOW-TO
# 삭제하고자하는 엑셀파일을 DST_PATH에 해당하는 경로에 위치
# DELETE_SHEET_NAME 에 해당하는 시트이름을 입력

# ### GLOBAL ########
OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

# ### VAR
SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

DELETE_SHEET_NAME = "Sheet1"

# ###

fileList = glob.glob(f'{SRC_PATH}*')

for file in fileList:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    fileAttr = file.split('\\')
    if file.find("~$") == -1:
        wb = excel.Workbooks.Open(Filename=file)
        sheets = wb.WorkSheets

        for sheet in sheets:
            if sheet.Name == DELETE_SHEET_NAME:
                print(f'Delete {sheet.Name}!')
                sheet.Delete()

        wb.SaveAs(f'{DST_PATH}{fileAttr[5]}')
        excel.Quit()