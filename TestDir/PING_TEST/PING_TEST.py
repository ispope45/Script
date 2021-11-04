from icmplib import ping


def ping_check(ip):
    try:
        response = ping(ip, count=10, timeout=0.038)
        print(response.is_alive)
        if response.is_alive:
            return True
        else:
            return False

    except Exception as e:
        print("Error : " + e)


if __name__ == "__main__":
    print(ping_check("8.8.8.8"))
