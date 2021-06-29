import socket

test_ip = '192.168.0.254'
test_port = range(50000, 50010)

for port in test_port:
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(0.5)
    # print(test_ip + " / " + str(port))
    result = sock.connect_ex((test_ip, port))
    sock.close()
    if result == 0:
        print(str(port) + " Port Opened")
    # else:
    #     print(str(port) + " Port Closed")

