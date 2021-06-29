import win32com.client
import os
# COMMENT #######
# 삼성전자 AP 설정파일 생성 툴(200814)

# ### GLOBAL ########
OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH + '\\Desktop\\'

SRC_PATH = '\\\\Filesvr\\인프라사업부\\02. 교육청 무선AP\\'
DST_PATH = '\\\\Filesvr\\인프라사업부\\02. 교육청 무선AP\\Config File\\'

SRC_FILE = '01. 교육청_무선AP_DB.xlsx'
SHEET_NAME = '무선AP_설치정보'
DST_FILE_NAME = 'config'
DST_FILE_EXT = '.txt'

# ### Row Control ######
# 시작과 끝라인 , 0으로 설정하면 맨첫줄(2) 부터 마지막줄까지
START_LINE = 7290
END_LINE = 7301

SEPARATE_FILE = True

# ### COL
# 순번/조직/이름/Hostname/IP/NETMASK/GATEWAY/APC_IP 순으로 배열
NO_COL = 2
ORG_NAME_COL = 7
NAME_COL = 8

HOSTNAME_COL = 9
IP_COL = 12
NETMASK_COL = 13
GATEWAY_COL = 14
SERVERIP_COL = 17

# 추가 멘트
# 설정값 맨 마지막줄에 추가
IS_EXTRA_COMMENT = False
EXTRA_COMMENT = 'BBBB\n\n'

# ###

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open(SRC_PATH + SRC_FILE)
ws = wb.WorkSheets(SHEET_NAME)

content = ''

if START_LINE == 0:
    START_LINE = 1

if END_LINE == 0:
    END_LINE = ws.UsedRange.Rows.Count

for row in range(START_LINE+2, END_LINE+3):
    if SEPARATE_FILE:
        fileName = str(int(ws.Cells(row, NO_COL).Value)) + "_" + str(ws.Cells(row, ORG_NAME_COL).Value) + "_" + \
                   str(ws.Cells(row, HOSTNAME_COL).Value) + DST_FILE_EXT

    else:
        fileName = DST_FILE_NAME + DST_FILE_EXT

    deviceIp = str(ws.Cells(row, IP_COL).Value)
    deviceNetmask = str(ws.Cells(row, NETMASK_COL).Value)
    deviceGateway = str(ws.Cells(row, GATEWAY_COL).Value)
    deviceSvrIp = str(ws.Cells(row, SERVERIP_COL).Value)
    deviceHostname = str(ws.Cells(row, HOSTNAME_COL).Value)

    script = (
            f'config interface address {deviceIp} {deviceNetmask} {deviceGateway}\n'
            f'config capwap apcIP {deviceSvrIp}\n'
            f'config capwap apName {deviceHostname}\n'
            f'show config interface summary\n\n')

    if IS_EXTRA_COMMENT:
        script = script + EXTRA_COMMENT

    if not SEPARATE_FILE:
        content = content + script
        content = content + '\n\n\n\n\n ################################################\n\n\n\n'

    if SEPARATE_FILE:
        f = open(DST_PATH + fileName, "w+")
        f.write(script)
        f.close()

if not SEPARATE_FILE:
    f = open(DST_PATH + fileName, "w+")
    f.write(content)
    f.close()

excel.Quit()