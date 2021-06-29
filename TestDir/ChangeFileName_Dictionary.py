import glob
import shutil
import os
import win32com.client

home_path = "C:\\Users\\Jungly\\Desktop\\"
src_path = "C:\\Users\\Jungly\\Desktop\\src\\"
dst_path = "C:\\Users\\Jungly\\Desktop\\dst\\"
xls_name = "33.xlsx"
'''
파일의 특정 문자열로 구분하여 파일 카피 및 정리
'''
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

wb = excel.Workbooks.Open(home_path + xls_name)
ws = wb.ActiveSheet

dic1 = {0: "None"}
key = []
for i in range(1, 48):
    dic1[int(ws.Cells(i, 1).Value)] = [int(ws.Cells(i, 2).Value), str(ws.Cells(i, 3).Value)]

f_list = glob.glob(src_path + "*")

for file in f_list:
    values = file.split('\\')
    key = values[5].split('.')

    fromFile = file
    toFile = dst_path + str(dic1[int(key[0])][0]) + "_" + dic1[int(key[0])][1] + ".jpg"

    print(fromFile + " to " + toFile)
    shutil.copy(fromFile, toFile)