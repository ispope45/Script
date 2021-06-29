import glob
import shutil
import os

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

PATH = "\\Ddnas\\인프라사업부_검수서류\\20200204_서울시교육청 망분리 및 고도화사업_2차(KT)"

TARGET_DIR = ['중부\\']


def main():

    fileList = glob.glob(f'{SRC_PATH}*\\*\\*')

    schNum, orgName, schName, fileName, fileExt, filePath = classify(fileList)

    copynpaste(schNum, orgName, schName, fileName, fileExt, filePath)

    # print(f'{schNum[0]}_{orgName[0]}_{schName[0]}_{fileName[0]}_{fileExt[0]}')


def copynpaste(schNum, orgName, schName, fileName, fileExt, filePath):

    for i in range(0, len(schNum)):
        dirNo = str(int(schNum[i]))
        dirOrg = str(orgName[i])
        dirName = str(schName[i])

        if not (os.path.isdir(DST_PATH + f'{dirNo}_{dirOrg}_{dirName}')):
            print(f'{dirNo}_{dirOrg}_{dirName}')
            os.mkdir(DST_PATH + f'{dirNo}_{dirOrg}_{dirName}')

    toFile = []
    fromFile = []

    a = ["현장실사점검표", "작업확인서", "망구성도", "준공서류"]

    for i in range(0, len(schNum)):
        toPath = f'{schNum[i]}_{orgName[i]}_{schName[i]}\\'
        if fileName[i].find("~$") == -1:

            if orgName[i].find("강동송파") != -1 or orgName[i].find("강서양천") != -1 \
                    or orgName[i].find("서부") != -1:

                if fileName[i].find("준공") != -1 or fileName[i].find("설치확인") != -1:
                    fromFile.append(filePath[i])
                    toFile.append(f'{DST_PATH}{toPath}{schName[i]}_{a[3]}.{fileExt[i]}')

                elif fileName[i].find("작업확인") != -1:
                    fromFile.append(filePath[i])
                    toFile.append(f'{DST_PATH}{toPath}{schName[i]}_{a[1]}.{fileExt[i]}')

                elif fileName[i].find("실사") != -1 or fileName[i].find("가.") != -1:
                    fromFile.append(filePath[i])
                    toFile.append(f'{DST_PATH}{toPath}{schName[i]}_{a[0]}.{fileExt[i]}')

                elif fileName[i].find("구성") != -1:
                    fromFile.append(filePath[i])
                    toFile.append(f'{DST_PATH}{toPath}{schName[i]}_{a[2]}.{fileExt[i]}')

            elif orgName[i].find("동작관악") != -1 or orgName[i].find("강남서초") != -1 \
                    or orgName[i].find("남부") != -1 or orgName[i].find("중부") != -1:

                if fileName[i].find("준공") != -1 or fileName[i].find("설치확인") != -1:
                    fromFile.append(filePath[i])
                    toFile.append(f'{DST_PATH}{toPath}{schName[i]}_{a[3]}.{fileExt[i]}')

                elif fileName[i].find("작업확인") != -1:
                    fromFile.append(filePath[i])
                    toFile.append(f'{DST_PATH}{toPath}{schName[i]}_{a[1]}.{fileExt[i]}')

                elif fileName[i].find("실사") != -1 or fileName[i].find("가.") != -1:
                    fromFile.append(filePath[i])
                    toFile.append(f'{DST_PATH}{toPath}{schName[i]}_{a[0]}.{fileExt[i]}')

                elif fileName[i].find("구성") != -1:
                    fromFile.append(filePath[i])
                    toFile.append(f'{DST_PATH}{toPath}{schName[i]}_{a[2]}.{fileExt[i]}')

        print(f'{len(fromFile)} :: {len(toFile)}')

    for j in range(0, len(fromFile)):
        print(f'{fromFile[j]} :: {toFile[j]}')
        if not (os.path.isfile(toFile[j])):
            # print(f'{fromFile[j]} :: {toFile[j]}')
            shutil.copy(fromFile[j], toFile[j])


def classify(fileList):
    schNum = []
    orgName = []
    schName = []
    fileName = []
    fileExt = []
    filePath = []

    for file in fileList:
        arr = file.split('\\')
        # arr.append()
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
        filePath.append(file)

    # print(len(schNum))
    return schNum, orgName, schName, fileName, fileExt, filePath


if __name__ == "__main__":
        main()
