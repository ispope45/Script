import socket
import os
import openpyxl

# GLOBAL
OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'
# SRC_FILE = HOME_PATH + '\\Desktop\\port.xlsx'
SRC_FILE = HOME_PATH + '\\Desktop\\Dev_List_20210706.xlsx'

# test_ip = ['192.168.0.254', '192.168.0.253', '192.168.0.252', '192.168.0.251', '192.168.0.250']
# test_port = [22, 80, 50005]
form_desc = '■ TCP Port Test\n' \
            'Read IP, Port number and TCP Port test\n' \
            'A:No(r) / B:Name(r) / C:IP(r) / D:Port(r) / E:Desc(w)\n'

START_LINE = 0
END_LINE = 5


# FUNCTION
def port_check(ip, port):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(0.5)
        result = sock.connect_ex((ip, port))
        sock.close()
    except socket.error as e:
        result = 1
        print("Error : " + e)

    if result == 0:
        print(str(ip) + ":" + str(port) + " Port Opened")
        return "Port Opened"
    else:
        print(str(ip) + ":" + str(port) + " Not Connected")
        return "Not Connected"


if __name__ == "__main__":
    print(form_desc)

    wb = openpyxl.load_workbook(SRC_FILE)
    ws = wb.active

    if START_LINE == 0:
        startRow = 2
    else:
        startRow = START_LINE

    if END_LINE == 0:
        totalRows = ws.max_row + 1
    else:
        totalRows = END_LINE

    for i in range(startRow, totalRows):
        row = str(i)

        val_ip = ws['C' + row].value
        val_port = ws['D' + row].value

        print(f"########### {i - 1} / {totalRows - 2} ###########")
        res = port_check(val_ip, val_port)
        ws['E' + row].value = res
        wb.save(filename=SRC_FILE)
