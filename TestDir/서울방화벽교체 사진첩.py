import win32com.client
from win32com.client import Dispatch
import shutil

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:\\Users\\Jungly\\Desktop\\4.xlsx')
ws = wb.ActiveSheet

arr = []

for i in range(2, 116):
    a = [ws.Cells(i, 1).Value, ws.Cells(i, 2).Value, ws.Cells(i, 3).Value, ws.Cells(i, 4).Value, ws.Cells(i, 5).Value, ws.Cells(i, 6).Value]
    # [순위, 지역, 학교명, 날짜, 담당자, 위치]
    arr.append(a)

# print(int(arr[0][0])) 순위 출력
# print(arr[0][7].strftime('%Y-%m-%d')) 날짜 출력

save_path = 'C:\\Users\\Jungly\\Desktop\\test1\\'
file_path = 'C:\\Users\\Jungly\\Desktop\\1.xlsx'


for a in arr:
    print(a[0])
    print(a[1])
    print(a[2])
    print(a[3])
    print(a[4])
    print(a[5])
    src = 'C:\\Users\\Jungly\\Desktop\\1.xlsx'
    dst = 'C:\\Users\\Jungly\\Desktop\\test1\\' + str(int(a[0])) + '_' + a[1] + '_' + a[2] + '_설치확인서.xlsx'

    #shutil.copy2(src, dst)

    xl = Dispatch("Excel.Application")
    xl.Visible = True

    wb1 = xl.Workbooks.Open(Filename=src)
    ws1 = wb1.Worksheets("완전포맷 사진첩(공통)")

    ws1.Cells(6, 4).Value = a[2]
    ws1.Cells(6, 9).Value = a[3].strftime('%Y-%m-%d')
    ws1.Cells(7, 4).Value = a[5]
    ws1.Cells(7, 9).Value = a[4]

    wb1.SaveAs(save_path + str(int(a[0])) + "_" + str(a[1]) + "_" + str(a[2]) + "_설치확인서.xlsx")
    wb1.Close(SaveChanges=True)

excel.Quit()