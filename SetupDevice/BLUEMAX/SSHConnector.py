import paramiko
import time


# GLOBAL
class SSHConnector():
    def __init__(self):
        self.preamble = ['y\n', 'config\n']
        self.hostname = ['hostname\n', 'set hostname\n', 'exit\n']
        self.interface = ['interface eth1\n', 'no ip add\n', 'ip add 100.100.100.100/24\n', 'exit\n', 'y\n']

    def main(self, sw_ip, sw_user, sw_pass, sw_port, command_set):
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
                    channel.send(line)

                    out_data, err_data = self.wait_streams(channel)
                    output.append(out_data)
                    print(out_data)


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

    @staticmethod
    def wait_streams(channel):
        time.sleep(1)
        out_data = ""
        err_data = ""
        while True:
            time.sleep(1.5)
            if channel.recv_ready():
                out_data += channel.recv(1000).decode('ascii')
                if channel.recv_stderr_ready():
                    err_data += channel.recv(1000).decode('ascii')
                if out_data.find("#") != -1 or out_data.find("[y|n]") != -1:
                    break

        return out_data, err_data

    def ssh_connect(self, con_info, ip_cfg):
        print(ip_cfg)
        self.main(con_info, "admin", "secui00@!", 22, ip_cfg)

