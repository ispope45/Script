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
                    # print(line)
                    channel.send(line)

                    out_data, err_data = self.wait_streams(channel)
                    output.append(out_data)
                    print(out_data)
                    # if out_data.find("Apply modifications") != -1:
                    #     while True:
                    #         time.sleep(2)
                    #         if channel.recv_ready():
                    #             if str(channel.recv(1000)).find("#") != -1:
                    #                 break
                    #             continue

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
        # if ha == "M":
        #     dev_list = ['eth1', 'eth2', 'eth3', 'eth9', 'eth9', 'eth11', 'eth12']
        #
        # elif ha == "A":
        #     dev_list = ['eth1', 'eth2', 'eth3', 'eth9', 'eth9', 'eth11', 'eth12']
        #
        # elif ha == "S":
        #     dev_list = ['eth1', 'eth2', 'eth3', 'eth9', 'eth9', 'eth12']
        # else:
        #     dev_list = []

        # cmd = []
        # host_cfg = self.hostname

        # host_cfg[1] = f'set {hostname}\n'
        # cmd.append(self.preamble)
        # cmd.append(ip_cfg)
        # cmd1.append(host_cfg)
        #
        # for i in range(0, len(dev_list)):
        #     if i == 3:
        #         cmd1.append([f'interface {dev_list[i]}\n', 'no ip add\n', 'no ip add\n',
        #                      f'ip add {ip1_list[i][0]}/{ip1_list[i][1]}\n',
        #                      f'ip add {ip1_list[i+1][0]}/{ip1_list[i+1][1]}\n', 'exit\n', 'y\n'])
        #     elif i == 4:
        #         continue
        #     else:
        #         cmd1.append([f'interface {dev_list[i]}\n', 'no ip add\n',
        #                      f'ip add {ip1_list[i][0]}/{ip1_list[i][1]}\n', 'exit\n', 'y\n'])

        print(ip_cfg)
        # self.main(con_info, "admin", "secui00@!", 22, ip_cfg)

