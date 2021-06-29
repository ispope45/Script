import win32com.client
import paramiko
from paramiko import AutoAddPolicy

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

wb = excel.Workbooks.Open('C:\\Users\\Jungle\\Desktop\\1.xlsx')
ws = wb.ActiveSheet
for i in range(2, 12):
    ssh_host = str(ws.Cells(i, 1).Value)
    ssh_user = str(ws.Cells(i, 2).Value)
    ssh_pass = str(ws.Cells(i, 3).Value)
    ssh_port = 22

    client = paramiko.SSHClient()
    client.load_system_host_keys()
    client.set_missing_host_key_policy(AutoAddPolicy())

    try:
        client.connect(hostname=ssh_host, username=ssh_user, password=ssh_pass, port=ssh_port)
        ws.Cells(i, 4).Value = "OK"
        print(ws.Cells(i, 1).Value + " OK")
    except Exception as e:
        ws.Cells(i, 4).Value = str(e)
        print(ws.Cells(i, 1).Value + " " + str(e))

    client.close()
