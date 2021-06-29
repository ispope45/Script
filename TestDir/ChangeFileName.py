import glob
import shutil
import os
import win32com.client

'''
파일의 특정 문자열로 구분하여 파일 카피 및 정리
'''
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

wb = excel.Workbooks.Open('C:\\Users\\Jungly\\Desktop\\3.xlsx')
ws = wb.ActiveSheet


#var 선언
name_list = []
name2_list = []
name3_list = []
aa = []
i = 0
count = 0

#파일 전체리스트
f_list = glob.glob("C:\\Users\\Jungly\\Desktop\\미래학교\\*")
aa = f_list[0].split('\\')

print(aa)
print(f_list)

for j in range(0, 35):
    shutil.copy(f_list[j], "C:\\Users\\Jungly\\Desktop\\TEST3\\" + str(int(ws.Cells(j + 2, 1).Value)) + "_" + str(ws.Cells(j + 2, 2).Value) + "_" + str(ws.Cells(j + 2, 3).Value) + ".jpg")

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

'''