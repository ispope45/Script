import glob
import shutil
import os

'''
파일의 특정 문자열로 구분하여 파일 카피 및 정리
'''

Path = "C:\\Users\\Jungly\\Desktop\\"

#var 선언
name_list = []
name2_list = []
name3_list = []
aa = []

from_filepath = []
to_filepath = []

i = 0
count = 0

#파일 전체리스트
f_list = glob.glob(Path + "src\\*")

aa = f_list[0].split('\\')
tmp = ""

for v in f_list:
    v2 = v.split('\\')

    '''
    if not (os.path.isdir("C:\\Users\\Jungly\\Desktop\\dst1\\" + v2[5])):
        print(v2[5])
        os.mkdir("C:\\Users\\Jungly\\Desktop\\dst1\\" + v2[5])
    '''

    if tmp == v2[5]:
        num = num + 1
    else:
        tmp = v2[5]
        num = 1

    shutil.copy(v, "C:\\Users\\Jungly\\Desktop\\dst1\\" + v2[5] + "_" + str(num) + "_" + v2[8])
    print(v2[5] + "_" + str(num))
    #shutil.copy(v, "C:\\Users\\Jungly\\Desktop\\dst\\" + v2[5] + "\\" + v2[8])

'''
for v in f_list:
    aa = v.split('\\')
    #print(aa)
    if aa[4].find("준공") != -1 and (aa[4].find("~$") == -1 and aa[4].find("xlsx") != -1):
        name_list.append(aa[4])
        name2_list.append(aa[4])
        name3_list.append("\\\\Ddnas\\학내망정비(2차_sk)\\" + aa[4] + "\\" + aa[4])
        #print(aa)

print(len(name2_list))
print(len(name_list))
print(len(name3_list))

for j in range(0, len(name2_list)):
    print(name3_list[j])
    print(name_list[j])
    print(name2_list[j])

if (len(name2_list) == len(name_list)) and (len(name2_list) == len(name3_list)):
    for j in range(0, len(name2_list)):
        shutil.copy(name3_list[j], "C:\\Users\\Jungle\\Desktop\\TEST3\\" + name_list[j] + "_SKB준공서류.xlsx")
        print(name3_list[j] + " to " + "C:\\Users\\Jungle\\Desktop\\TEST3\\" + name_list[j] + "\\" + name2_list[j])
'''
'''
if (len(name2_list) == len(name_list)) and (len(name2_list) == len(name3_list)):
    for j in range(0, len(name2_list)):
        if not(os.path.isdir("D:\\Desktop\\TEST3\\" + name_list[j])):
            os.mkdir("D:\\Desktop\\TEST3\\" + name_list[j])
        shutil.copy(name3_list[j], "D:\\Desktop\\TEST3\\" + name_list[j] + "\\" + name2_list[j])
        print(name3_list[j] + " to " + "C:\\Users\\Jungle\\Desktop\\TEST3\\" + name_list[j] + "\\" + name2_list[j])
'''
# 서류 분류
'''
for v in f_list:
    aa = v.split('\\')
    name_list.append(aa[5])
    ext = aa[6].split('.')

    if aa[6].find("준공") != -1 and (aa[6].find("~$") == -1):
        from_filepath.append(v)
        to_filepath.append("C:\\Users\\Jungly\\Desktop\\KT준공서류\\" + aa[5] + "_KT준공서류.xlsx")
    elif aa[6].find("망구성") != -1 and (aa[6].find("~$") == -1):
        from_filepath.append(v)
        to_filepath.append("C:\\Users\\Jungly\\Desktop\\망구성도\\" + aa[5] + "_망구성도.pptx")
    elif aa[6].find("현장실사") != -1 and (aa[6].find("~$") == -1):
        from_filepath.append(v)
        to_filepath.append("C:\\Users\\Jungly\\Desktop\\현장실사점검표\\" + aa[5] + "_현장실사점검표.xlsx")
    elif aa[6].find("작업") != -1 and (aa[6].find("~$") == -1):
        from_filepath.append(v)
        to_filepath.append("C:\\Users\\Jungly\\Desktop\\작업확인서\\" + aa[5] + "_작업확인서." + ext[-1])
    else:
        count = count + 1
        name2_list.append("")

if len(from_filepath) == len(to_filepath):
    loop_size = len(from_filepath)

for num in range(loop_size):
    print(from_filepath[num])
    print(to_filepath[num])
    shutil.copy(from_filepath[num], to_filepath[num])
'''
'''
for v in f_list:
    aa = v.split('\\')
    name_list.append(aa[7])

    bb = aa[7].split('_')
    print(bb)
    print(bb[2])

    if aa[8].find("작업") != -1 and (aa[8].find("~$") == -1):
        name2_list.append(aa[8])
    elif aa[8].find("준공") != -1 and (aa[8].find("~$") == -1):
        name2_list.append(aa[8])
    elif aa[8].find(".jpg") != -1 and (aa[8].find("~$") == -1):
        count = count + 1
        name2_list.append("")
    else:
        count = count + 1
        name2_list.append("")

# 무결성 검증
print(len(name_list))
print(len(name2_list))
print(len(f_list))
print(count)

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