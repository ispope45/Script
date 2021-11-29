import paramiko
import time
import os
import openpyxl
import socket

# GLOBAL

cmd1 = ['enable\n', 'show interface status\n', ' ', ' ', 'show power inline port-status\n', ' ', ' ']
cmd2 = ['enable\n', 'show power inline port-status\n']

cmd = ['y\n', 'config\n', 'interface eth1\n', 'no ip add\n', 'ip add 100.100.100.100/24\n', 'exit\n', 'y\n']


def main(sw_ip, sw_user, sw_pass, sw_port, command):
    host = sw_ip
    username = sw_user
    password = sw_pass

    output = list()

    try:
        conn = paramiko.SSHClient()
        conn.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        conn.connect(host, username=username, password=password, port=sw_port, timeout=10)
        channel = conn.invoke_shell()
        time.sleep(2)

        for line in command:
            channel.send(line)
            out_data, err_data = wait_streams(channel)
            print(out_data)
            # tmp = out_data.split("\\r\\n")
            # print(out_data)

            output.append(out_data)

        return output

    except Exception as e:
        if "port" in str(e):
            return "Port Error"
        if "WinError 10060" in str(e):
            return "Connection Error"
        print(e)
        return "Error"

    finally:
        if conn is not None:
            conn.close()


def wait_streams(channel):
    time.sleep(1)
    out_data = ""
    err_data = ""
    print(channel.recv_ready())
    while channel.recv_ready():
        time.sleep(1)
        out_data += channel.recv(1000).decode('ascii')
    while channel.recv_stderr_ready():
        err_data += str(channel.recv_stderr(1000))

    return out_data, err_data


if __name__ == "__main__":
    res = main("192.168.10.10", "admin", "secui00@!", 22, cmd)
    print(res)
