import paramiko
import time
import os
import openpyxl
import socket

# GLOBAL
preamble = ['config\n']

cmd1 = [preamble,
        ['interface eth1\n', 'no ip add\n', 'ip add 100.100.100.100/24\n', 'exit\n'],
        ['interface eth2\n', 'no ip add\n', 'ip add 200.200.200.200/24\n', 'exit\n'],
        ['interface eth3\n', 'no ip add\n', 'ip add 130.100.100.100/24\n', 'exit\n'],
        ['interface eth9\n', 'no ip add\n', 'ip add 230.200.200.200/24\n', 'exit\n'],
        ['interface eth11\n', 'no ip add\n', 'ip add 240.200.200.200/24\n', 'exit\n'],
        ['interface eth12\n', 'no ip add\n', 'ip add 250.200.200.200/24\n', 'exit\n']]

cmd2 = [preamble,
        ['interface eth1\n', 'no ip add\n', 'ip add 100.100.100.100/24\n', 'exit\n', 'y\n'],
        ['interface eth2\n', 'no ip add\n', 'ip add 200.200.200.200/24\n', 'exit\n', 'y\n'],
        ['interface eth3\n', 'no ip add\n', 'ip add 100.100.100.100/24\n', 'exit\n', 'y\n'],
        ['interface eth9\n', 'no ip add\n', 'ip add 200.200.200.200/24\n', 'exit\n', 'y\n'],
        ['interface eth11\n', 'no ip add\n', 'ip add 200.200.200.200/24\n', 'exit\n', 'y\n'],
        ['interface eth12\n', 'no ip add\n', 'ip add 200.200.200.200/24\n', 'exit\n', 'y\n']]


def main(sw_ip, sw_user, sw_pass, sw_port, command_set):
    host = sw_ip
    username = sw_user
    password = sw_pass

    output = list()
    try:
        conn = paramiko.SSHClient()
        conn.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        conn.connect(host, username=username, password=password, port=sw_port, timeout=100)
        channel = conn.invoke_shell()
        time.sleep(3)

        for command in command_set:
            for line in command:
                res = ""
                while True:
                    if res.find("#") != -1:
                        break

                    if res.find("[y|n]") != -1:
                        res += send(channel, 'y')

                    if channel.recv_ready():
                        if str(channel.recv(1000)).find("#") != -1:
                            break
                        continue
                    time.sleep(2)
                    res = ""
                    res += send(channel, '\n')

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
    print(out_data)
    print(err_data)
    return out_data


def wait_streams(channel):
    time.sleep(1)
    out_data = ""
    err_data = ""
    while True:
        time.sleep(1)
        if channel.recv_ready():
            out_data += channel.recv(1000).decode('ascii')
            if channel.recv_stderr_ready():
                err_data += channel.recv(1000).decode('ascii')
            break

    return out_data, err_data


if __name__ == "__main__":
    res = main("192.168.10.10", "admin", "secui00@!", 22, cmd1)
    # res = main("192.168.10.11", "admin", "secui00@!", 22, cmd1)
    print(res)
