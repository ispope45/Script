import win32com.client
from win32com.client import Dispatch
import datetime

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:\\Users\\Jungle\\Desktop\\4.xlsx')
ws = wb.ActiveSheet

arr = []

for i in range(2, 48):
    a = [ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, ws.Cells(i, 5).Value, ws.Cells(i, 6).Value, ws.Cells(i, 7).Value]
    # [순위, 지역, 학교명, 담당자, AP1위치, AP2위치, PoE위치]
    arr.append(a)

# print(int(arr[0][0])) 순위 출력
# print(arr[0][7].strftime('%Y-%m-%d')) 날짜 출력

save_path = 'C:\\Users\\Jungle\\Desktop\\test2\\'
file_path = 'C:\\Users\\Jungle\\Desktop\\2.xlsx'

for a in arr:
    xl = Dispatch("Excel.Application")
    xl.Visible = True

    wb1 = xl.Workbooks.Open(Filename=file_path)
    ws1 = wb1.ActiveSheet

    ws1.Cells(5, 4).Value = a[2]
    ws1.Cells(6, 4).Value = a[3]
    ws1.Cells(9, 2).Value = a[4]
    ws1.Cells(18, 2).Value = a[5]
    ws1.Cells(27, 2).Value = a[6]

#xl = Dispatch("Excel.Application")
#xl.Visible = True
#
#wb1 = xl.Workbooks.Open(Filename=file_path)
#ws1 = wb1.ActiveSheet
#
#print(arr[0][3])
#print(arr[0][7].strftime('%Y-%m-%d'))
#print(arr[0][4] + "\n" + arr[0][5])
#print(arr[0][6])

#ws1.Cells(6, 3).Value = arr[0][3]
#ws1.Cells(6, 7).Value = arr[0][7].strftime('%Y-%m-%d')
#ws1.Cells(13, 7).Value = arr[0][4] + "\n" + arr[0][5]
#ws1.Cells(14, 7).Value = arr[0][6]

    wb1.SaveAs(save_path + str(int(a[0])) + "_" + str(a[1]) + "_" + str(a[2]) + "_설치사진첩.xlsx")
    wb1.Save()
    wb1.Close(SaveChanges=True)

excel.Quit()