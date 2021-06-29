import glob
import shutil
import os

'''
파일의 특정 문자열로 구분하여 파일 카피 및 정리
'''

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
f_list = glob.glob("C:\\Users\\Jungly\\Desktop\\src\\*\\*\\*")
aa = f_list[0].split('\\')

# ['C:', 'Users', 'Jungly', 'Desktop', 'src', '강남서초', '100_강남서초_경기여자고등학교', '100_강남서초_경기여자고등학교_KT준공서류.xlsx']
# print(aa)
# print(f_list)

# 서류 분류
for v in f_list:
    aa = v.split('\\')

    ext = aa[7].split('.')

    if aa[7].find(".ppt") != -1 and (aa[7].find("~$") == -1):
        from_filepath.append(v)
        name_list.append(aa[6])
        to_filepath.append("C:\\Users\\Jungly\\Desktop\\dst\\" + aa[6] + "_망구성도.pptx")
    # else:
    #     count = count + 1
    #     to_filepath.append("")

    print(len(from_filepath))
    print(len(to_filepath))


if len(from_filepath) == len(to_filepath):
    i = len(to_filepath)
    filename = to_filepath[0].split("\\")[5]

    for j in range(0, i):
        filename = to_filepath[j].split("\\")[5]
        if not(os.path.isdir("C:\\Users\\Jungly\\Desktop\\dst\\" + name_list[j])):
            print(name_list[j])
            os.mkdir("C:\\Users\\Jungly\\Desktop\\dst\\" + name_list[j])
        if name_list[j]:
            print(from_filepath[j] + " to " + "C:\\Users\\Jungly\\Desktop\\dst\\" + name_list[j] + "\\" + filename)
            shutil.copy(from_filepath[j],  "C:\\Users\\Jungly\\Desktop\\dst\\" + name_list[j] + "\\" + filename)
