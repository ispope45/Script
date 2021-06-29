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

# 파일 전체리스트
# f_list = glob.glob("C:\\Users\\Jungly\\Desktop\\서울시교육청_학내망정비사업_검수서류(동작관악)_최종\\*\\*")

# 테스트
f_list = glob.glob("C:\\Users\\Jungly\\Desktop\\TEST1\\*\\*\\*")

aa = f_list[0].split('\\')

print(len(aa))
print(f_list)

# 서류 분류
for v in f_list:
    aa = v.split('\\')

    if aa[len(aa)-1].find("현장실사") != -1 and (aa[len(aa)-1].find("~$") == -1):
        name_list.append(v)
        name2_list.append(aa[len(aa)-1])

'''
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open(name_list[0])
ws = wb.ActiveSheet

isDev = False

# 장비현황 (51~99 , 9)
for j in range(51, 99):
    if ws.Cells(j, 9).Value:
        if (ws.Cells(j, 9).Value.find("공인") != -1 or ws.Cells(j, 9).Value.find("서비스") != -1) or isDev:
            # print(ws.Cells(i, 1).Value)
            isDev = True

print(isDev)
# 사용망현황 (19~23, 1)
# Eth3 포함항목 명칭 DMZ망으로 명칭 변경
for i1 in range(18, 22):
    if ws.Cells(i1, 1).Value.find("th3") != -1:
        if isDev:
            ws.Cells(i1, 1).Value = "DMZ망(Eth3)"
        else:
            ws.Cells(i1, 1).Value = None
    else:
        print("No Change : ")


# 빈항목이 있고, 아래항목이 있으면 아래항목을 위로 올림
for i2 in range(19, 22):
    # print(ws.Cells(i2, 1).Value + " ::: " + ws.Cells(i2+1, 1).Value)
    if not(ws.Cells(i2, 1).Value) and ws.Cells(i2+1, 1).Value:
        ws.Cells(i2, 1).Value = ws.Cells(i2+1, 1).Value
        ws.Cells(i2, 5).Value = ws.Cells(i2+1, 5).Value
        ws.Cells(i2, 7).Value = ws.Cells(i2+1, 7).Value
        ws.Cells(i2+1, 1).Value = None
        ws.Cells(i2+1, 5).Value = None
        ws.Cells(i2+1, 7).Value = None
        print("Test1")
        break

# Eth0 포함항목 명칭 기타서비스망으로 명칭 변경
for i3 in range(18, 22):
    if not(ws.Cells(i3, 1).Value):
        if ws.Cells(i3, 1).Value.find("기타") != -1:
            ws.Cells(i3, 1).Value = "기타서비스망(Eth0)"
        else:
            print("No Change : ")

wb.Close(SaveChanges=True)
excel.Quit()
'''

for path in name_list:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    wb = excel.Workbooks.Open(path)
    ws = wb.ActiveSheet

    isDev = False

    print(path)

    # 장비현황 (51~99 , 9)
    for j in range(51, 99):
        if ws.Cells(j, 9).Value:
            if (ws.Cells(j, 9).Value.find("공인") != -1 or ws.Cells(j, 9).Value.find("서비스") != -1) or isDev:
                # print(ws.Cells(i, 1).Value)
                isDev = True

    print("공인망 " + str(isDev))
    # 사용망현황 (19~23, 1)
    # Eth3 포함항목 명칭 DMZ망으로 명칭 변경
    for i1 in range(18, 22):
        if (ws.Cells(i1, 1).Value):
            if ws.Cells(i1, 1).Value.find("th3") != -1:
                if isDev:
                    ws.Cells(i1, 1).Value = "DMZ망(Eth3)"
                else:
                    ws.Cells(i1, 1).Value = None

    # 빈항목이 있고, 아래항목이 있으면 아래항목을 위로 올림
    for i2 in range(19, 22):
        # print(ws.Cells(i2, 1).Value + " ::: " + ws.Cells(i2+1, 1).Value)
        if not(ws.Cells(i2, 1).Value) and ws.Cells(i2+1, 1).Value:
            ws.Cells(i2, 1).Value = ws.Cells(i2+1, 1).Value
            ws.Cells(i2, 5).Value = ws.Cells(i2+1, 5).Value
            ws.Cells(i2, 7).Value = ws.Cells(i2+1, 7).Value
            ws.Cells(i2+1, 1).Value = None
            ws.Cells(i2+1, 5).Value = None
            ws.Cells(i2+1, 7).Value = None
            break

    # Eth0 포함항목 명칭 기타서비스망으로 명칭 변경
    for i3 in range(18, 22):
        if ws.Cells(i3, 1).Value:
            if ws.Cells(i3, 1).Value.find("기타") != -1:
                ws.Cells(i3, 1).Value = "기타서비스망(Eth0)"
                print("기타망 수정")

    print("")
    wb.Close(SaveChanges=True)
    excel.Quit()



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