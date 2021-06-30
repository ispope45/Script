import paramiko
import time
import win32com.client
import os
import csv


OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
CUR_PATH = os.getcwd()
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH + '\\Desktop\\'
SRC_FILE = CUR_PATH + '\\Python_diag.xlsx'
#SRC_FILE = '\\\\Filesvr\\인프라사업부\\95. 정기점검\\고려대학교부속병원_구로병원\\Python_diag.xlsx'
DST_PATH = CUR_PATH + '\\'


def main(sw_ip, sw_user, sw_pass):
    host = sw_ip
    username = sw_user
    password = sw_pass

    output = list()
    command = list()
    command.append('enable\n')
    command.append('show syslog non-volatile tail\n')
    command.append('show temperature\n')

    try:
        conn = paramiko.SSHClient()
        conn.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        conn.connect(host, username=username, password=password)
        channel = conn.invoke_shell()
        time.sleep(3)

        for line in command:
            channel.send(line)
            out_data, err_data = wait_streams(channel)

            a = out_data.replace("\\r\\n", "\n").replace("\\r", "")
            b = a.replace('b"', "").replace("b'", "").replace(" '", "").replace('"', "")

            output.append(b)

        return output

    except Exception as e:
        print(e)
        return "Error"

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
    print(CUR_PATH)
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(SRC_FILE)
    ws = wb.ActiveSheet

    for row in range(2, ws.UsedRange.Rows.Count+1):
        f = open(DST_PATH + f'{str(int(ws.Cells(row, 1).Value))}_{ws.Cells(row, 2).Value}.txt', "w+")
        sw_ip = ws.Cells(row, 3).Value
        sw_user = ws.Cells(row, 4).Value
        sw_pass = ws.Cells(row, 5).Value
        e = main(sw_ip, sw_user, sw_pass)
        for line in e:
            print(line)
            f.write(line)
        f.close()
    excel.Quit()
