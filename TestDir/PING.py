from icmplib import ping


def ping_check(ip):
    try:
        response = ping(ip, count=1)
        if response.is_alive:
            return True
        else:
            return False

    except Exception as e:
        print("Error : " + e)


host = '8.8.8.8'
res = ping_check(host)
print(res)