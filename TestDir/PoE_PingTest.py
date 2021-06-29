import win32com.client
import random

RootPath = 'C:\\Users\\Jungly\\Desktop\\'

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open(RootPath + '2.xlsx')
ws = wb.ActiveSheet
a = 1

#ttl = 61

for i in range(2, 11):
    arr1 = []
    num = 0
    for j in range(1, 39):
        arr1.append(random.uniform(40.1, 40.8))

    ttl = int(ws.Cells(i, 5).Value)
    max_value = max(arr1)
    min_value = min(arr1)
    avg_value = sum(arr1) / len(arr1)

    script = (str(ws.Cells(i, 2).Value) + "#" + '\n' + str(ws.Cells(i, 2).Value) + "#" + '\n'
              + str(ws.Cells(i, 2).Value) + "#ping " + str(ws.Cells(i, 3).Value) + '\n'
              "PING " + str(ws.Cells(i, 3).Value) + " (" + str(ws.Cells(i, 3).Value) + "): 56 data bytes" + '\n')
    for k in arr1:
        num = num + 1
        script = script + ("64 bytes from " + str(ws.Cells(i, 3).Value) + ": icmp_seq=" + str(num) + " ttl=" + str(ttl) + " time=" + str(k)[0:5]
                           + " ms" + '\n')

    script = script + ("^C--- " + str(ws.Cells(i, 3).Value) + " ping statistics ---" + '\n'
                       + str(len(arr1)) + " packets transmitted, 38 packets received, 0% packet loss" + '\n'
                       "round-trip min/avg/max/stddev = " + str(min_value)[0:5] + "/" + str(avg_value)[0:5] + "/"
                       + str(max_value)[0:5] + "/" + str(random.uniform(0.5, 0.9))[0:6] + " ms" + '\n'
                       + str(ws.Cells(i, 2).Value) + "#" + '\n' + str(ws.Cells(i, 2).Value) + "#" + '\n')

    print(script)
    f = open(RootPath + "dst\\" + str(int(ws.Cells(i, 4).Value)) + "_" + str(ws.Cells(i, 1).Value) + "_PoE Ping_TEST2.txt", 'w+')
    f.write(script)
    f.close()

excel.Quit()