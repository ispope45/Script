import glob
from win32com.client import Dispatch
import win32com.client
import shutil
import os

'''
여러경로의 특정파일들을 찾아 엑셀파일 특정내용을 수정
'''

#var 선언
name_list = []
name2_list = []
name3_list = []
aa = []
i = 0
count = 0

#파일 전체리스트
f_list = glob.glob("C:\\Users\\Jungly\\Desktop\\서울시교육청_학내망정비사업_검수서류(동작관악)_최종\\*\\*")
aa = f_list[0].split('\\')

print(aa)
print(f_list)

# 서류 분류
for v in f_list:
    aa = v.split('\\')
    #name_list.append(aa[6])

    #bb = aa[7].split('_')
    #print(bb)
    #print(bb[2])

    if aa[6].find("현장실사") != -1 and (aa[6].find("~$") == -1):
        name_list.append(v)
        name2_list.append(aa[6])

#print(name_list)
#print(name2_list)
'''
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open(name_list[0])
ws = wb.ActiveSheet
print(name_list[0])
print(ws.Cells(21, 1).Value)
# ws.Cells(21, 1).Value = "기타서비스망(Eth3)"

wb.Close(SaveChanges=True)
excel.Quit()

'''

cnt = 0

for path in name_list:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Open(path)
    ws = wb.ActiveSheet

    if ws.Cells(19, 1).Value.find("기타") != -1:
        print(ws.Cells(19, 1).Value)
        ws.Cells(19, 1).Value = "기타서비스망(Eth0)"
        cnt = cnt + 1
    elif ws.Cells(20, 1).Value.find("기타") != -1:
        print(ws.Cells(20, 1).Value)
        ws.Cells(20, 1).Value = "기타서비스망(Eth0)"
        cnt = cnt + 1
    elif ws.Cells(21, 1).Value.find("기타") != -1:
        print(ws.Cells(21, 1).Value)
        ws.Cells(21, 1).Value = "기타서비스망(Eth0)"
        cnt = cnt + 1
    elif ws.Cells(22, 1).Value.find("기타") != -1:
        print(ws.Cells(22, 1).Value)
        ws.Cells(22, 1).Value = "기타서비스망(Eth0)"
        cnt = cnt + 1
    elif ws.Cells(23, 1).Value.find("기타") != -1:
        print(ws.Cells(23, 1).Value)
        ws.Cells(23, 1).Value = "기타서비스망(Eth0)"
        cnt = cnt + 1
    wb.Close(SaveChanges=True)
    excel.Quit()

print(cnt)

# 무결성 검증
'''
if (len(name2_list) == len(name_list)) and (len(f_list) == len(name2_list)) and (len(f_list) == len(name_list)):
    i = len(name2_list)
    print(i)
    print(name2_list)
    for j in range(0, i):
        if not(os.path.isdir("D:\\Desktop\\TEST3\\" + name_list[j])):
            os.mkdir("D:\\Desktop\\TEST3\\" + name_list[j])
        if name2_list[j]:
            print(f_list[j] + " to " + name_list[j])
        #    print(f_list[j] + " to " + "D:\\Desktop\\TEST3\\" + name_list[j] + "\\" + name2_list[j] )
            shutil.copy(f_list[j], "D:\\Desktop\\TEST3\\" + name_list[j] + "\\" + name2_list[j])

else:
    print("Error: List indistinct Check for filename")
'''