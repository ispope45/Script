import paramiko
import time
import win32com.client
import os

# 안랩방화벽 경유 SSH 접근 및 Config Tool

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH
SRC_FILE = HOME_PATH + '\\Desktop\\src.xlsx'


def main(fw_ip, fw_user, fw_pass, sw_ip, sw_user, sw_pass, sw_gw, nms_ip):
    host = fw_ip
    username = fw_user
    password = fw_pass

    command = list()
    command.append('rhgoelqjrld\n')
    command.append('ghkdrmaenRjql\n')
    command.append(f'ssh {sw_user}@{sw_ip}\n -p 2004')
    command.append(f'{sw_pass}\n')
    command.append('enable\n')
    command.append('configure terminal\n')
    command.append(f'ip route 0.0.0.0/0 {sw_gw}\n')
    command.append('snmp-server group schoolnet3 v3 priv\n')
    command.append('snmp-server user schoolnet3 schoolnet3 v3 auth md5 Sen16701396! Priv des Sen16701396!\n')
    command.append(f'snmp-server host {nms_ip} version 3 priv schoolnet3\n')
    command.append('snmp-server enable traps\n')
    command.append('end\n')
    command.append('write memory\n')
    command.append('y\n')

    try:
        conn = paramiko.SSHClient()
        conn.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        conn.connect(host, username=username, password=password)
        channel = conn.invoke_shell()
        time.sleep(3)
        for line in command:
            channel.send(line)
            out_data, err_data = wait_streams(channel)
            if out_data.find("Are you sure you want to continue connecting (yes/no)? ") != -1:
                channel.send("yes\n")
                out_data, err_data = wait_streams(channel)
            elif out_data.find("Connection refused") != -1:
                return "SSH Port Error"
            elif out_data.find("No route to host") != -1:
                return "Host IP Error"

            print(out_data.replace("\\r\\n", "\n").replace("b'", ">>>"))

        return "OK"

    except Exception as e:
        print(e)
        return "FW Connection Error"

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
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    wb = excel.Workbooks.Open(SRC_FILE)
    ws = wb.ActiveSheet

    for row in range(2, ws.UsedRange.Rows.Count+1):
        fw_ip = ws.Cells(row, 4).Value
        fw_user = ws.Cells(row, 5).Value
        fw_pass = ws.Cells(row, 6).Value
        sw_ip = ws.Cells(row, 7).Value
        sw_user = ws.Cells(row, 8).Value
        sw_pass = ws.Cells(row, 9).Value
        sw_gw = ws.Cells(row, 10).Value
        nms_ip = ws.Cells(row, 11).Value
        e = main(fw_ip, fw_user, fw_pass, sw_ip, sw_user, sw_pass, sw_gw, nms_ip)
        if e:
            ws.Cells(row, 12).Value = e
            wb.Save()
    excel.Quit()
