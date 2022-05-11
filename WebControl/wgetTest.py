import wget
import math
import os
import win32com.client

BASEURL = "http://edu.im.neostep.kr/upload/school/"

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']

HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH + "\\Desktop\\"
TARGET_DIR = "E:\\result\\"
# TARGET_DIR = "result\\"
TARGET_FILE = HOME_PATH + "backup.xlsx"
BASEPATH = os.path.dirname(os.path.realpath(__file__))
# TARGET_DIR = BASEPATH + "\\result\\"
# TARGET_FILE = BASEPATH + "\\backup.xlsx"
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

START_LINE = 0
END_LINE = 0


def bar_custom(current, total, width=80):
    width=30
    avail_dots = width-2
    shaded_dots = int(math.floor(float(current) / total * avail_dots))
    percent_bar = '[' + '■'*shaded_dots + ' '*(avail_dots-shaded_dots) + ']'
    progress = "%d%% %s [%d / %d]" % (current / total * 100, percent_bar, current, total)
    return progress


def download(url, out_path):
    try:
        wget.download(url, out=out_path, bar=bar_custom)
        print(" Ok")
    except:
        print(url + " Error : 404 Not Exist")
        f = open(TARGET_DIR + "\\log.txt", "a+")
        f.write(out_path + " Error : 404 Not Exist\n")
        f.close()


if __name__ == "__main__":

    print(os.path.realpath(__file__))
    print(os.path.dirname(os.path.realpath(__file__)))

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(TARGET_FILE)
    ws = wb.ActiveSheet

    if START_LINE == 0:
        startRow = 2
    else:
        startRow = START_LINE

    if END_LINE == 0:
        totalRows = ws.UsedRange.Rows.Count + 1
    else:
        totalRows = END_LINE

    # url = BASEURL + "7/report.pdf"
    if not os.path.isdir(TARGET_DIR):
        os.makedirs(TARGET_DIR, exist_ok=True)
        print("Create Dir " + TARGET_DIR)
    for row in range(startRow, totalRows):
        No = str(int(ws.Cells(row, Column["COL_A"]).value))
        oriNo = str(int(ws.Cells(row, Column["COL_B"]).value))
        Name = ws.Cells(row, Column["COL_C"]).value

        fileName = oriNo + "_" + Name + ".pdf"
        reqUrl = BASEURL + No + "/report.pdf"

        # print(TARGET_DIR + fileName)
        print(reqUrl)

        resultFile = TARGET_DIR + fileName
        download(reqUrl, resultFile)

    excel.Quit()
