import win32com.client
import glob
import os

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']
HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH

# SRC_PATH = HOME_PATH + '\\Desktop\\src\\'
SRC_PATH = '\\\\192.168.0.50\\인프라사업부\\15. 서울특별시교육청 학교 무선통신장비 도입 및 설치\\31. 산출물\\장비사진'
DST_PATH = HOME_PATH + '\\Desktop\\dst\\'

TARGET_FILE = '\\\\192.168.0.50\\인프라사업부\\15. 서울특별시교육청 학교 무선통신장비 도입 및 설치\\★ 사진검토2.xlsx'

START_LINE = 0
END_LINE = 0

# NO	학교	PoE	집선	AP	전면	후면	전면	후면	전면	후면	신호측정
Column = {"COL_A": 1,
          "COL_B": 2,
          "COL_C": 3,
          "COL_D": 4,
          "COL_E": 5,
          "COL_F": 6,
          "COL_G": 7,
          "COL_H": 8,
          "COL_I": 9,
          "COL_J": 10,
          "COL_K": 11,
          "COL_L": 12,
          "COL_M": 13,
          "COL_N": 14,
          "COL_O": 15,
          "COL_P": 16,
          "COL_Q": 17,
          "COL_R": 18,
          "COL_S": 19,
          "COL_T": 20}
ROW_NUM = 3

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(TARGET_FILE)
ws = wb.ActiveSheet

if START_LINE == 0:
    startRow = 2
else:
    startRow = START_LINE

if END_LINE == 0:
    totalRows = ws.UsedRange.Rows.Count
else:
    totalRows = END_LINE

firstRow = startRow
lastRow = totalRows


print(SRC_PATH)
files = glob.glob(f'{SRC_PATH}\\*\\*')
print(files)
# Initialize
ws.Range('F3:M386').value = None

# for file in files:
#     # print(file)
#     arr1 = file.split('\\')
#     # print(arr1)  # ['', '', '192.168.0.50', '인프라사업부', '15. 서울특별시교육청 학교 무선통신장비 도입 및 설치', '31. 산출물', '장비사진', 'PoE', '870_서울오정초등학교_PoE04_후면.jpg']
#
#     device = arr1[7]
#     fileName = arr1[8]
#     if fileName.find("_") != -1:
#         value = fileName.replace(".jpg", "").split("_")
#         value.append(device)
#
#         resArr.append(value)
#
# print(resArr)
# print(len(resArr))
count = []
resArr = []

for i in range(3, totalRows):
    # print(ws.Cells(i, Column["COL_A"]).value)
    count.append(int(ws.Cells(i, Column["COL_A"]).value))

for file in files:
    # print(file)
    arr1 = file.split('\\')
    # print(arr1)  # ['', '', '192.168.0.50', '인프라사업부', '15. 서울특별시교육청 학교 무선통신장비 도입 및 설치', '31. 산출물', '장비사진', 'PoE', '870_서울오정초등학교_PoE04_후면.jpg']

    device = arr1[7]
    fileName = arr1[8]
    if fileName.find("_") != -1:
        value = fileName.replace(".jpg", "").split("_")
        value.append(device)
        resArr.append(value)

print(resArr)
for resVal in resArr:
    rowIdx = count.index(int(resVal[0])) + 3

    print(resVal)
    if resVal[3] == "설치확인서":
        if resVal[2] == "설치확인서":
            colIdx = Column["COL_M"]
    elif resVal[4] == "AP":
        if resVal[3] == "전면":
            colIdx = Column["COL_J"]
        elif resVal[3] == "후면":
            colIdx = Column["COL_K"]
        elif resVal[3] == "신호측정":
            colIdx = Column["COL_L"]
    elif resVal[4] == "PoE":
        if resVal[3] == "전면":
            colIdx = Column["COL_F"]
        elif resVal[3] == "후면":
            colIdx = Column["COL_G"]
    elif resVal[4] == "BB":
        if resVal[3] == "전면":
            colIdx = Column["COL_H"]
        elif resVal[3] == "후면":
            colIdx = Column["COL_I"]
    # elif resVal[4] == "배치도":
    #     paintCnt = paintCnt + 1

    # print(ws.Cells(rowIdx, colIdx).value)

    if ws.Cells(rowIdx, colIdx).value is None:
        ws.Cells(rowIdx, colIdx).value = 1
    else:
        ws.Cells(rowIdx, colIdx).value = int(ws.Cells(rowIdx, colIdx).value) + 1

# print(count.index(10))
#
# print(resArr)
# print(len(resArr))
# print(count)

wb.Save()
excel.Quit()
