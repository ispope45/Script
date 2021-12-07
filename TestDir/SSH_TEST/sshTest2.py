import paramiko
import time
import os
import openpyxl
import socket

# GLOBAL
cmd1 = ['show_all\n']


def main(sw_ip, sw_user, sw_pass, sw_port, command):
    host = sw_ip
    username = sw_user
    password = sw_pass

    output = list()
    try:
        conn = paramiko.SSHClient()
        conn.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        conn.connect(host, username=username, password=password, port=sw_port, timeout=100)
        channel = conn.invoke_shell()
        time.sleep(1)

        for line in command:
            res = ""
            res += send(channel, line)
            output.append(res)

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


def send(channel, cmd):
    channel.send(cmd)
    out_data, err_data = wait_streams(channel)
    a = out_data.replace("\\r\\n", "\n").replace("\\r", "")
    b = a.replace('b"', "").replace("b'", "").replace('"', "")
    return b


def wait_streams(channel):
    time.sleep(1)
    out_data = ""
    err_data = ""
    while True:
        time.sleep(1)
        if channel.recv_ready():
            out_data += str(channel.recv(100000))
            if channel.recv_stderr_ready():
                err_data += str(channel.recv(100000))
            break

    return out_data, err_data


if __name__ == "__main__":
    res = main("192.168.0.254", "manager", "qwe123!@#", 22, cmd1)
    f = open("TEST.txt", "w")
    for val in res:
        f.write(val)
