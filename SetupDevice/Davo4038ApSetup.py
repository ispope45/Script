import win32com.client
import os

# COMMENT #######
# 다보링크 4038이전 AP 설정파일 생성 툴(200818)
# 1_강남서초_개원중학교.txt

# ### GLOBAL ########
OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

TEST = 1

SRC_FILE = 'backup.xlsx'
DST_FILE_NAME = 'config'
DST_FILE_EXT = '.txt'

# ### Row Control ######
# 시작과 끝라인 , 0으로 설정하면 맨첫줄(2) 부터 마지막줄까지
START_LINE = 0
END_LINE = 0

SEPARATE_FILE = True

# ### COL
# 순번/조직/이름/Hostname/IP/NETMASK/GATEWAY/APC_IP 순으로 배열
NO_COL = 1
ORG_NAME_COL = 2
NAME_COL = 3

HOSTNAME_COL = 4
IP_COL = 5
NETMASK_COL = 6
GATEWAY_COL = 7
SERVERIP_COL = 8

DNS1 = '168.126.63.1'
DNS2 = '168.126.63.2'
# 추가 멘트
# 설정값 맨 마지막줄에 추가
IS_EXTRA_COMMENT = False
EXTRA_COMMENT = 'BBBB\n\n'

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
    if SEPARATE_FILE:
        fileName = str(int(ws.Cells(row, NO_COL).Value)) + "_" + str(ws.Cells(row, ORG_NAME_COL).Value) + "_" + \
                   str(ws.Cells(row, NAME_COL).Value) + DST_FILE_EXT

    else:
        fileName = DST_FILE_NAME + DST_FILE_EXT

    deviceIp = str(ws.Cells(row, IP_COL).Value)
    deviceNetmask = str(ws.Cells(row, NETMASK_COL).Value)
    deviceGateway = str(ws.Cells(row, GATEWAY_COL).Value)
    deviceSvrIp = str(ws.Cells(row, SERVERIP_COL).Value)
    deviceHostname = str(ws.Cells(row, HOSTNAME_COL).Value)
    deviceDns1 = DNS1
    deviceDns2 = DNS2

    script = (
            f'config\n'
            f'basic wan\n'
            f'static\n'
            f'{deviceIp}\n'
            f'{deviceNetmask}\n'
            f'{deviceGateway}\n'
            f'{deviceDns1}\n'
            f'{deviceDns2}\n'
            f'y\n'
            f'basic ap_name\n'
            f'{deviceHostname}\n'
            f'basic wtp\n'
            f'{deviceSvrIp}\n'
            f'basic ap_mode\n'
            f'1\n'
            f'apply\n\n')

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