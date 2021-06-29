import os
import glob

f_list = glob.glob("C:\\Users\\Jungle\\Desktop\\TEST3\\*")
aa = f_list[0].split('\\')
#var 선언
fileList = []
aa = []
i = 0

for v in f_list:
    aa = v.split('\\')
    print(aa)
    fileList.append(aa[5])

for v1 in fileList:
    tmp = v1.split('_')
    print(tmp[0])
