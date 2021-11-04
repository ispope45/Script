from telnetlib import Telnet

user = "root"
passwd = "frontier"
host = "218.237.161.106"
port = "22223"
proto = "TELNET"


def encoder(text):
    res = text.encode('utf-8') + b'\n'
    return res

def conn_tel(user, passwd, host, port, proto):
    tn = Telnet(host, port=port, timeout=5)
    with Telnet(host, port) as tn:
        r = [[b"login: ", encoder(user)], [b"Password: ", encoder(passwd)], [b"Guro_5F_24P_02#", encoder("show logging back")]]
        for val in r:
            print(val)
            response = tn.read_until(val[0])
            print(response)
            tn.write(val[1])
            print(tn.read_all().decode('utf-8'))

        tn.write(b"exit\n")
        tn.close()
        # tn.interact()
        # tn.close()
    #
    # r = [[b"login: ", encoder(user)], [b"Password: ", encoder(passwd)], [b"Guro_5F_24P_02#", encoder("show logging back")]]
    #
    # for val in r:
    #     print(val)
    #     response = tn.read_until(val[0])
    #     print(response)
    #     tn.write(val[1])
    #
    # tn.write(b"exit\n")
    # print(tn.read_all().decode('utf-8'))
    # r = tn.read_until(b"login: ")
    #
    #
    # print(r)
    # w = tn.write(user.encode('utf-8') + b'\n')
    # print(w)
    # r = tn.read_until(b"Password: ")
    # print(r)
    # tn.write(passwd.encode('utf-8') + b'\n')
    # print(tn)
    # tn.write(b"exit\n")
    #
    # print(tn.read_all().decode('utf-8'))


if __name__ == "__main__":
    conn_tel(user, passwd, host, port, proto)