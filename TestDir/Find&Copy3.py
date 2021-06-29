import glob
import os

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

PATH = "\\Ddnas\\인프라사업부_검수서류\\20200204_서울시교육청 망분리 및 고도화사업_2차(KT)"

TARGET_DIR = ['강남서초\\', '강동송파\\', '강서양천\\', '동작관악\\', '남부\\', '중부\\', '서부\\']


def main(target):
    print(target)
    fileList = glob.glob(f'{SRC_PATH}{target}*\\*')
    schNum, orgName, schName, fileName, fileExt = classify(fileList)
    # print(f'{schNum}')
    # print(f'{fileName}')
    # makedir(schNum, orgName, schName)
    # print(len(schNum))
    # print(len(schName))
    # print(len(orgName))
    # print(len(fileName))
    # print(len(fileExt))


def makedir(schNum, orgName, schName):
    for i in range(0, len(schNum)):
        fileNo = str(int(schNum[i]))
        fileOrg = str(orgName[i])
        fileName = str(schName[i])

        if not (os.path.isdir(DST_PATH + f'{fileNo}_{fileOrg}_{fileName}')):
            print(f'{fileNo}_{fileOrg}_{fileName}')
            # os.mkdir(DST_PATH + f'{fileNo}_{fileOrg}_{fileName}')


def copynpaste():
    print(0)


def classify(fileList):
    keyword = ['준공', '설치확인', '실사', '작업확인', '구성']

    schNum = []
    orgName = []
    schName = []
    fileName = []
    fileExt = []

    for file in fileList:
        # if (key in file for key in keyword):
        #     print("Catch")
        for key in keyword:
            isFile = False
            if file.find("$") == -1 and file.find(key) != -1:
                isFile = True
                break

        if isFile:
            arr = file.split('\\')
            arr.append()
            # ['C:', 'Users', 'Jungly', 'Desktop', 'src', '서부', '9_서부_서울상신초등학교', '라. 서울상신초등학교_준공서류(SK).xlsx']
            schAttr = arr[6].split('_')
            # ['9', '서부', '서울상신초등학교']
            fileAttr = arr[7].split('.')
            # ['라', ' 서울상신초등학교_준공서류(SK)', 'xlsx']
            schNum.append(int(schAttr[0]))
            orgName.append(schAttr[1])
            schName.append(schAttr[2])
            fileName.append(arr[7])
            fileExt.append(fileAttr[-1])

    print(len(schNum))
    return schNum, orgName, schName, fileName, fileExt


if __name__ == "__main__":
    for target in TARGET_DIR:
        main(target)
