import win32com.client
import glob
import os

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

TARGET_FILE = HOME_PATH + '\\Desktop\\fw1.xlsx'

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

EXTRA_COMMENT = (
    "object ip_address ipv4_address add '172.28.228.70_WNMS' '0' '172.28.228.70' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.71_SODE' '0' '172.28.228.71' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.101_NiaAdm' '0' '172.28.228.101' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.102_NiaAdm' '0' '172.28.228.102' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.103_NiaAdm' '0' '172.28.228.103' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.104_NiaAdm' '0' '172.28.228.104' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.105_NiaAdm' '0' '172.28.228.105' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.106_NiaAdm' '0' '172.28.228.106' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.107_NiaAdm' '0' '172.28.228.107' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.108_NiaAdm' '0' '172.28.228.108' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.109_NiaAdm' '0' '172.28.228.109' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.110_NiaAdm' '0' '172.28.228.110' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '192.168.72.82_NiaEMS1' '0' '192.168.72.82' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '192.168.72.83_NiaEMS2' '0' '192.168.72.83' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.11_NiaAuthSvr1' '0' '172.28.228.11' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.21_NiaAuthSvr2' '0' '172.28.228.21' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.77_wNMS1' '0' '172.28.228.77' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.78_wNMS2' '0' '172.28.228.78' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.79_wNMS3' '0' '172.28.228.79' '32' 'eth7' '1' '' ''\n"
    "object ip_address ipv4_address add '172.28.228.80_wNMS4' '0' '172.28.228.80' '32' 'eth7' '1' '' ''\n\n"
    
    "# Svc object\n"
    "object service service add 'TCP_80' '6' '1-65535' '80-80' '3600' '' ''\n"
    "object service service add 'TCP_88' '6' '1-65535' '88-88' '3600' '' ''\n"
    "object service service add 'TCP_161' '6' '1-65535' '161-161' '3600' '' ''\n"
    "object service service add 'TCP_6000' '6' '1-65535' '6000-6000' '3600' '' ''\n"
    "object service service add 'TCP_7557' '6' '1-65535' '7557-7557' '3600' '' ''\n"
    "object service service add 'TCP_7547' '6' '1-65535' '7547-7547' '3600' '' ''\n"
    "object service service add 'TCP_8080' '6' '1-65535' '8080-8080' '3600' '' ''\n"
    "object service service add 'TCP_8443' '6' '1-65535' '8443-8443' '3600' '' ''\n"
    "object service service add 'TCP_8445' '6' '1-65535' '8445-8445' '3600' '' ''\n"
    "object service service add 'TCP_8822' '6' '1-65535' '8822-8822' '3600' '' ''\n"
    "object service service add 'TCP_8888' '6' '1-65535' '8888-8888' '3600' '' ''\n"
    "object service service add 'TCP_9445' '6' '1-65535' '9445-9445' '3600' '' ''\n"
    "object service service add 'TCP_9992' '6' '1-65535' '9992-9992' '3600' '' ''\n"   
    "object service service add 'UDP_ALL' '17' '1-65535' '1-65535' '30' '' ''\n"
    "object service service add 'UDP_161' '17' '1-65535' '161-161' '30' '' ''\n"
    "object service service add 'UDP_162' '17' '1-65535' '162-162' '30' '' ''\n"
    "object service service add 'UDP_514' '17' '1-65535' '514-514' '30' '' ''\n"
    "object service service add 'UDP_1812' '17' '1-65535' '1812-1812' '30' '' ''\n"
    "object service service add 'UDP_1813' '17' '1-65535' '1813-1813' '30' '' ''\n"
    "object service service add 'UDP_2055' '17' '1-65535' '2055-2055' '30' '' ''\n\n\n"
    
    "# Policy\n"
    "policy firewall ipv4 add '0' '1' '172.28.228.70_WNMS;' 'WirelessNet' 'TCP_7547;TCP_8445;TCP_9445;UDP_ALL' '2' 'always' '1' 'TRSvr_log' '' '0' ''\n"
    "policy firewall ipv4 add '0' '1' '172.28.228.71_SODE' 'WirelessNet' 'TCP_8445;TCP_9445' '2' 'always' '1' 'TRSvr2_log' '' '0' ''\n"
    "policy firewall ipv4 add '0' '1' '172.28.228.77_wNMS1;172.28.228.78_wNMS2;172.28.228.79_wNMS3;172.28.228.80_wNMS4' 'WirelessNet' 'PING;UDP_161;UDP_162;UDP_514;TCP_161' '2' 'always' '1' 'NiaWnms_log' '' '0' ''\n"
    "policy firewall ipv4 add '0' '1' '192.168.72.82_NiaEMS1;192.168.72.83_NiaEMS2' 'WirelessNet' 'PING;UDP_161;UDP_162;UDP_514;UDP_2055;TCP_161' '2' 'always' '1' 'NiaEMS_log' '' '0' ''\n"
    "policy firewall ipv4 add '0' '1' '172.28.228.11_NiaAuthSvr1;172.28.228.21_NiaAuthSvr2' 'WirelessNet' 'UDP_1812;UDP_1813' '2' 'always' '1' 'NiaAuthSvr_log' '' '0' ''\n"
    "policy firewall ipv4 add '0' '1' '172.28.228.101_NiaAdm;172.28.228.102_NiaAdm;172.28.228.103_NiaAdm;172.28.228.104_NiaAdm;172.28.228.105_NiaAdm;172.28.228.106_NiaAdm;172.28.228.107_NiaAdm;172.28.228.108_NiaAdm;172.28.228.109_NiaAdm;172.28.228.110_NiaAdm' 'WirelessNet' 'PING;TCP_8822;TCP_88;TCP_6000;TCP_8080;TCP_9992;TCP_8888' '2' 'always' '1' 'NiaAdm_log' '' '0' ''\n"

    )

# EXTRA_COMMENT = ''
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

    if str(ws.Cells(i, Column["COL_Q"]).value) != "공동회선":  # 공동회선(병설유치원)
        # script = (
        #     f'en\n'
        #     f'config terminal\n'
        #     f'interface vlan 1\n'
        #     f'ip address {ws.Cells(i, Column["COL_S"]).value}{str(ap_ip)}/{str(int(ws.Cells(i, Column["COL_P"]).value))}\n'
        #     f'exit\n\n'
        #     f'hostname {ws.Cells(i, Column["COL_H"]).value}_BB\n'
        #     f'ip route 0.0.0.0/0 {ws.Cells(i, Column["COL_R"]).value}\n\n'
        #     f'service ssh\n'
        #     f'ip ssh port 9992\n\n'
        #     f'end\n'
        #     f'write memory\n'
        #     f'y\n\n')
        nid = ws.Cells(i, Column["COL_S"]).value
        subnet = ws.Cells(i, Column["COL_T"]).value
        name = ws.Cells(i, Column["COL_T"]).value
        script = (
            f"# IP object\n"
            f"object ip_address ipv4_address add 'WirelessNet' '1' '{nid}' '{int(subnet)}' 'all' '1' '' ''\n\n"
        )

        data = script + EXTRA_COMMENT
        # print(string)
        f = open(DST_PATH + str(int(ws.Cells(i, Column["COL_A"]).value)) + "_" + ws.Cells(i, Column["COL_H"]).value + ".txt", "w+")
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
