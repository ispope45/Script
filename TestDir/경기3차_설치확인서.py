import win32com.client
from win32com.client import Dispatch
import datetime

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:\\Users\\Jungly\\Desktop\\22.xlsx')
ws = wb.ActiveSheet

arr = []

for i in range(2, 35):
    a = [ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, ws.Cells(i, 5).Value]
    # [순위, 지역, 학교명, 담당자, ap1, ap2, poe1, 날짜]
    arr.append(a)

# print(int(arr[0][0])) 순위 출력
# print(arr[0][7].strftime('%Y-%m-%d')) 날짜 출력

save_path = 'C:\\Users\\Jungly\\Desktop\\test1\\'
file_path = 'C:\\Users\\Jungly\\Desktop\\223.xlsx'

for a in arr:
    xl = Dispatch("Excel.Application")
    xl.Visible = True

    wb1 = xl.Workbooks.Open(Filename=file_path)
    ws1 = wb1.ActiveSheet

    ws1.Cells(6, 3).Value = a[2]
    ws1.Cells(13, 7).Value = a[4]
    ws1.Cells(14, 7).Value = a[3]
#    ws1.Cells(27, 3).Value = a[3]


    wb1.SaveAs(save_path + str(int(a[0])) + "_" + str(a[1]) + "_" + str(a[2]) + "_설치확인서.xlsx")
    wb1.Save()
    wb1.Close(SaveChanges=True)

excel.Quit()