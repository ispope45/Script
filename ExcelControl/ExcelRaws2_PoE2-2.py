import win32com.client
import glob
import os

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

TARGET_FILE = HOME_PATH + '\\Desktop\\poe4.xlsx'

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

# EXTRA_COMMENT = (
#     'service ssh\n'
#     'ip ssh port 9992\n\n'
#     'end\n'
#     'write memory\n'
#     'y\n\n'
#     '---------------- 추가설정(직접입력) ---------------- \n\n'
#     'conf terminal\n'
#     'snmp-server community RO\n'
#     'aiedu\n'
#     'aiedu\n'
#     'snmp-server community RW\n'
#     'niaedu\n'
#     'niaedu\n\n'
#     'write memory\n'
#     'y\n\n'
#     '---------------- 계정정보(직접입력) ---------------- \n\n'
#     '기본값 : root // [패스워드없음]\n'
#     '** root 입력시 계정재생성 메세지출력됨\n'
#     '변경값 : admin // frontier1!\n\n')
EXTRA_COMMENT = (
    'service ssh\n'
    'ip ssh port 9992\n\n'
    'end\n'
    'write memory\n'
    'y\n\n')
ROW_NUM = 2
POE_IP = 220

LOG_FILE = "excel.log"
RES_FILE = "20210127_POE.csv"

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


print(str(firstRow) + " : " + str(lastRow))
f = open(DST_PATH + RES_FILE, "w+")
# attribute = "no,class,name,attr1,hostname,networkId,subnetMask,gatewayIp,apIp\n"
# f.write(attribute)

totalNum = 0

for i in range(firstRow, lastRow+1):
    print(f'{str(i-1)}/{str(lastRow-1)}')
    data = ""
    # print(ws.Cells(i, Column["COL_D"]))
    ap_count = int(ws.Cells(i, Column["COL_G"]).value)
    ap_ip = 221
    sb_ap_num = 0
    ap_num = 0

    if int(ws.Cells(i, Column["COL_H"]).value) == 1:  # 공동회선(병설유치원)
        sb_ap_num = sb_ap_num + 1
        totalNum = totalNum + 1

        script = (
            f'en\n'
            f'config terminal\n'
            f'interface vlan 1\n'
            f'ip address {ws.Cells(i, Column["COL_N"]).value}220/{str(int(ws.Cells(i, Column["COL_Q"]).value))}\n'
            f'exit\n\n'
            f'hostname {ws.Cells(i, Column["COL_M"]).value}_BB\n'
            f'ip route 0.0.0.0/0 {ws.Cells(i, Column["COL_P"]).value}\n\n')

        data = script + EXTRA_COMMENT
        # print(string)
        f = open(DST_PATH + str(int(ws.Cells(i, Column["COL_A"]).value)) + "_" + ws.Cells(i, Column["COL_D"]).value + "_BB.txt", "w+")
        f.write(data)
        f.close()

    for j in range(0, ap_count):
        ap_num = ap_num + 1
        totalNum = totalNum + 1
        if ap_num < 10:
            number = "0" + str(ap_num)
        else:
            number = str(ap_num)
        script = (
            f'en\n'
            f'config terminal\n'
            f'interface vlan 1\n'
            f'ip address {ws.Cells(i, Column["COL_N"]).value}{str(ap_ip)}/{str(int(ws.Cells(i, Column["COL_Q"]).value))}\n'
            f'exit\n\n'
            f'hostname {ws.Cells(i, Column["COL_M"]).value}_PoE{str(number)}\n'
            f'ip route 0.0.0.0/0 {ws.Cells(i, Column["COL_P"]).value}\n\n')
        ap_ip = ap_ip + 1
        data = script + EXTRA_COMMENT
        # print(string)
        f = open(DST_PATH + str(int(ws.Cells(i, Column["COL_A"]).value)) + "_" + ws.Cells(i, Column["COL_D"]).value + "_PoE" + str(number) + ".txt", "w+")
        f.write(data)
        f.close()

excel.Quit()

# for i in range(firstRow, lastRow+1):
#     ap_count = int(ws.Cells(i, Column["COL_I"]).value)
#     apCount = 0
#     for j in range(rows, rows + ap_count):
#         rows = rows + 1
#         apCount = apCount + 1


#
# rows = 2
#
# ws2.Cells(1, Column["COL_A"]).value = "순번"
# ws2.Cells(1, Column["COL_B"]).value = "집선청"
# ws2.Cells(1, Column["COL_C"]).value = "학교명"
# ws2.Cells(1, Column["COL_D"]).value = "Hostname"
# ws2.Cells(1, Column["COL_E"]).value = "설치업체"
# ws2.Cells(1, Column["COL_F"]).value = "IP"
# ws2.Cells(1, Column["COL_G"]).value = "Subnet"
# ws2.Cells(1, Column["COL_H"]).value = "GW"
#
#
# for i in range(firstRow, lastRow+1):
#     ap_count = int(ws.Cells(i, Column["COL_I"]).value)
#     apCount = 0
#     for j in range(rows, rows + ap_count):
#         rows = rows + 1
#         apCount = apCount + 1
#         ws2.Cells(j, Column["COL_A"]).value = ws.Cells(i, Column["COL_A"]).value
#         ws2.Cells(j, Column["COL_B"]).value = ws.Cells(i, Column["COL_B"]).value
#         ws2.Cells(j, Column["COL_C"]).value = ws.Cells(i, Column["COL_C"]).value
#         ws2.Cells(j, Column["COL_D"]).value = ws.Cells(i, Column["COL_D"]).value
#         ws2.Cells(j, COL_E).value = ws.Cells(i, COL_E).value
#         ws2.Cells(j, COL_F).value = ws.Cells(i, COL_H).value
#         ws2.Cells(j, COL_G).value = ws.Cells(i, COL_L).value
#         ws2.Cells(j, COL_H).value = apCount
#
# print("Total AP Counts : " + count)
# wb2.SaveAs(DST_PATH + "res.xlsx")

#
