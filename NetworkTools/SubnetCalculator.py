import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

Column = {"COL_A": 1,
          "COL_B": 2,
          "COL_C": 3,
          "COL_D": 4,
          "COL_E": 5,
          "COL_F": 6,
          "COL_G": 7,
          "COL_H": 8,
          "COL_I": 9,
          "COL_J": 10,
          "COL_K": 11,
          "COL_L": 12,
          "COL_M": 13,
          "COL_N": 14,
          "COL_O": 15,
          "COL_P": 16,
          "COL_Q": 17,
          "COL_R": 18,
          "COL_S": 19,
          "COL_T": 20}

START_LINE = 0
END_LINE = 0

SRC_FILE = 'C:\\Users\\Jungly\\Desktop\\backup.xlsx'


def sub_cal(value):
    netmask = value.split("/")[1]
    ipOctet = value.split("/")[0].split(".")
    subnet = []
    sub1 = int(int(netmask) / 8)
    sub2 = int(netmask) % 8

    resNet = []
    resBrd = []
    resStartHost = []
    resEndHost = []

    for i in range(0, sub1):
        subnet.append(255)

    if len(subnet) <= 4:
        i = len(subnet)
        subnet.append(0)
        if sub2 == 0:
            subnet[i] = 0
        elif sub2 == 1:
            subnet[i] = 128
        elif sub2 == 2:
            subnet[i] = 192
        elif sub2 == 3:
            subnet[i] = 224
        elif sub2 == 4:
            subnet[i] = 240
        elif sub2 == 5:
            subnet[i] = 248
        elif sub2 == 6:
            subnet[i] = 252
        elif sub2 == 7:
            subnet[i] = 254

    for i in range(len(subnet), 4):
        subnet.append(0)

    for i in range(0, 4):
        resNet.append(int(ipOctet[i]) & subnet[i])
        resBrd.append(int(ipOctet[i]) | (~subnet[i] + 256))
        if i == 3:
            resStartHost.append((int(ipOctet[i]) & subnet[i]) + 1)
            resEndHost.append((int(ipOctet[i]) | (~subnet[i] + 256)) - 1)
        else:
            resStartHost.append(int(ipOctet[i]) & subnet[i])
            resEndHost.append(int(ipOctet[i]) | (~subnet[i] + 256))

    return resNet, resBrd, resStartHost, resEndHost


def ip_merge(ip_set):
    res = str(ip_set[0])
    res = res + "." + str(ip_set[1])
    res = res + "." + str(ip_set[2])
    res = res + "." + str(ip_set[3])

    return res


if __name__ == "__main__":

    wb = excel.Workbooks.Open(SRC_FILE)
    ws = wb.ActiveSheet

    if START_LINE == 0:
        startRow = 2
    else:
        startRow = START_LINE

    if END_LINE == 0:
        totalRows = ws.UsedRange.Rows.Count + 1
    else:
        totalRows = END_LINE

    for row in range(startRow, totalRows):
        _NETWORK = ws.Cells(row, Column["COL_D"]).value
        _NETMASK = _NETWORK.split("/")[1]
        _IP_OCTET = _NETWORK.split("/")[0].split(".")

        for ip in _IP_OCTET:
            if int(ip) > 256:
                print("error")
            elif int(ip) < 0:
                print("error2")

        resNet, resBrd, resStartHost, resEndHost = sub_cal(_NETWORK)

        ws.Cells(row, Column["COL_H"]).value = ip_merge(resNet)
        ws.Cells(row, Column["COL_I"]).value = ip_merge(resBrd)
        ws.Cells(row, Column["COL_J"]).value = ip_merge(resStartHost)
        ws.Cells(row, Column["COL_K"]).value = ip_merge(resEndHost)
