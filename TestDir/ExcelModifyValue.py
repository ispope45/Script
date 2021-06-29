from win32com.client import Dispatch
import win32com.client
import random
import glob
'''
엑셀 특정값을 RANDOM으로 수정 하여 저장
'''
f_list = glob.glob("D:\\Desktop\\excel\\*")
'''
a = ['2', '18'], ['2', '19'], ['2', '20'], ['2', '21'], ['2', '22'], ['2', '25'], ['2', '26'], ['2', '27'], \
    ['2', '28'], ['3', '4'], ['3', '5'], ['3', '6'], ['3', '7'], ['3', '8'], ['3', '11'], ['3', '12'], ['3', '13'], \
    ['3', '14'], ['3', '15'], ['3', '18'], ['3', '19'], ['3', '20'], ['3', '21'], ['3', '22'], ['3', '25'], \
    ['3', '26'], ['3', '27'], ['3', '28'], ['3', '29'], ['4', '1'], ['4', '2'], ['4', '3'], ['4', '4'], ['4', '5'], \
    ['4', '8'], ['4', '9'], ['4', '10'], ['4', '11'], ['4', '12']
'''

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('D:\\Desktop\\3.xlsx')
ws = wb.ActiveSheet

arr = []
k = 0

for i in range(2,48):
    j = ws.Cells(i, 4).Value
    arr.append(j)

for f in f_list:
    xl = Dispatch("Excel.Application")
    xl.Visible = True

    path1 = f
    aa = f.split('\\')
    print(path1)
    print(aa)

    bb = aa[3].split('_')
    print(int(bb[0]))

    wb1 = xl.Workbooks.Open(Filename=path1)
    ws1 = wb1.ActiveSheet

    ws1.Cells(9, 2).Value = arr[int(bb[0])-1]
    k = k + 1

    wb1.Close(SaveChanges=True)

    xl.Quit()
