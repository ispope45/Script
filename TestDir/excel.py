import win32com.client

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:\\Users\\Jungle\\Desktop\\test.xlsx')
ws = wb.ActiveSheet
for i in range(1,1241):
    script = ('config interface address ' + ws.Cells(i,2).Value + ' ' + ws.Cells(i,3).Value + ' ' + ws.Cells(i,4).Value + '\n'
    'config capwap apcIP ' + ws.Cells(i,6).Value + '\n'
    'config capwap apName ' + ws.Cells(i,5).Value + '\n'
    'show config interface summary')
    f = open("C:\\Users\\Jungle\\Desktop\\NIA\\"+ ws.Cells(i,1).Value + ".txt", 'w')
    f.write(script)
    f.close()

excel.Quit()