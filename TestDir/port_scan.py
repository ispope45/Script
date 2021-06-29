import socket
import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

wb = excel.Workbooks.Open('C:\\Users\\Jungly\\Desktop\\1.xlsx')
ws = wb.ActiveSheet

for i in range(2, 12):
    try:
        s = socket.socket()
        s.connect((ws.Cells(i, 1).Value, 22))
        banner = s.recv(1024)

        ws.Cells(i, 5).Value = "[" + ws.Cells(i, 1).Value + "] " + str(banner) + " is Open"
        s.close()
    except Exception as e:
        ws.Cells(i, 5).Value = "[" + ws.Cells(i, 1).Value + "] Fail " + str(e)
