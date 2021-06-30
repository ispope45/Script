import paramiko
import time
import win32com.client
import os

# 서울시교육청 무선스위치 SNMP 통합설정

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH
SRC_FILE = HOME_PATH + '\\Desktop\\12.xlsx'

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

START_LINE = 0
END_LINE = 0

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False


def main(s_ip, s_user, s_pass, s_port):
    host = s_ip
    username = s_user
    password = s_pass
    port = s_port

    command = list()
    command.append('enable\n')
    command.append('configure terminal\n')
    command.append('snmp-server community rw 1\n')
    command.append('test1\n')  # SNMP Community String
    command.append('test1\n')

    conn = paramiko.SSHClient()

    try:

        conn.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        conn.connect(host, port=port, username=username, password=password)
        channel = conn.invoke_shell()
        time.sleep(4)
        for line in command:
            channel.send(line)
            out_data, err_data = wait_streams(channel)

            print(out_data.replace("\\r\\n", "\n").replace("b'", ">>>"))

        return "OK"

    except Exception as e:
        print(e)
        return str(e)

    finally:
        if conn is not None:
            conn.close()


def wait_streams(channel):
    time.sleep(1)
    out_data = ""
    err_data = ""

    while channel.recv_ready():
        out_data += str(channel.recv(1000))
    while channel.recv_stderr_ready():
        err_data += str(channel.recv_stderr(1000))

    return out_data, err_data


if __name__ == "__main__":
    wb = excel.Workbooks.Open(SRC_FILE)
    ws = wb.ActiveSheet
    if START_LINE == 0:
        startRow = 2
    else:
        startRow = START_LINE

    if END_LINE == 0:
        totalRows = ws.UsedRange.Rows.Count + 1
    else:
        totalRows = END_LINE

    for row in range(startRow, totalRows):
        s_ip = ws.Cells(row, Column['COL_B']).Value
        s_user = ws.Cells(row, Column['COL_C']).Value
        s_pass = ws.Cells(row, Column['COL_D']).Value
        s_port = int(ws.Cells(row, Column['COL_E']).Value)

        e = main(s_ip, s_user, s_pass, s_port)
        if e:
            ws.Cells(row, Column['COL_F']).Value = e
            wb.Save()
    excel.Quit()
