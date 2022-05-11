import win32com.client
import os

# COMMENT #######
# 유비쿼스 PoE 설정파일 생성 툴(200818)
# NIA 사업용 Config 파일임

# SAMPLE(8-24LINE)
''' 

en
config terminal

interface vlan 1
ip address 10.62.206.250/23
exit

hostname Seoul_garak.es_poe01
ip route 0.0.0.0/0 10.62.207.254
end

write memory
y

'''

# ### GLOBAL ########
OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

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

# 추가 멘트
# 설정값 맨 마지막줄에 추가
IS_EXTRA_COMMENT = True
EXTRA_COMMENT = (
    'service ssh\n'
    'ip ssh port 9992\n\n'
    'end\n'
    'write memory\n'
    'y\n\n'
    '---------------- 추가설정(직접입력) ---------------- \n\n'
    'conf terminal\n'
    'snmp-server community RO\n'
    'aiedu\n'
    'aiedu\n'
    'snmp-server community RW\n'
    'niaedu\n'
    'niaedu\n\n'
    'write memory\n'
    'y\n\n'
    '---------------- 계정정보(직접입력) ---------------- \n\n'
    '기본값 : root // [패스워드없음]\n'
    '** root 입력시 계정재생성 메세지출력됨\n'
    '변경값 : admin // frontier1!\n\n')

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

    script = (
            f'en\n'
            f'config terminal\n'
            f'interface vlan 1\n'
            f'ip address {deviceIp}{deviceNetmask}\n'
            f'exit\n\n'
            f'hostname {deviceHostname}\n'
            f'ip route 0.0.0.0/0 {deviceGateway}\n'
            f'end\n\n'
            f'write memory\n'
            f'y\n\n')

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
