import wget
import math
import os
import time
import random

BASEURL = "http://m.ajit.ac.kr/my/lec/"

OS_HOME_DRIVE = os.environ['HOMEDRIVE']
OS_HOME_PATH = os.environ['HOMEPATH']

HOME_PATH = OS_HOME_DRIVE + OS_HOME_PATH + "\\Desktop\\"
TARGET_DIR = HOME_PATH + "result\\"
TARGET_DIR = "E:\\result\\"
# TARGET_FILE = HOME_PATH + "src.xlsx"
# BASEPATH = os.path.dirname(os.path.realpath(__file__))
# TARGET_DIR = BASEPATH + "\\result\\"
# TARGET_FILE = BASEPATH + "\\src.xlsx"

CONTENTS_ID = [114]
# END_LECTURE_ID = [140, 83, 78, 79]
END_LECTURE_ID = 150
START_LINE = 0
END_LINE = 0
ERROR_CNT = 0


def bar_custom(current, total, width=80):
    width = 30
    avail_dots = width - 2
    shaded_dots = int(math.floor(float(current) / total * avail_dots))
    percent_bar = '[' + '■'*shaded_dots + ' '*(avail_dots-shaded_dots) + ']'
    progress = "%d%% %s [%d / %d]" % (current / total * 100, percent_bar, current, total)
    return progress


def download(url, out_path, error_cnt):
    try:
        wget.download(url, out=out_path, bar=bar_custom)
        print(" Ok")
        return int(error_cnt)
    except Exception as e:
        print(url + " Error : 404 Not Exist")
        f = open(TARGET_DIR + "\\log.txt", "a+")
        f.write(out_path + f" Error : 404 Not Exist : {e}\n")
        f.close()
        return int(error_cnt + 1)


if __name__ == "__main__":
    # url = BASEURL + "7/report.pdf"
    if not os.path.isdir(TARGET_DIR):
        os.makedirs(TARGET_DIR, exist_ok=True)
        print("Create Dir " + TARGET_DIR)

    for i in range(0, len(CONTENTS_ID)):
        for Lid in range(1, END_LECTURE_ID):
            reqUrl = BASEURL + str(CONTENTS_ID[i]) + "/" + str(Lid) + ".mp4"
            print(BASEURL + str(CONTENTS_ID[i]) + "/" + str(Lid) + ".mp4")
            fileName = TARGET_DIR + str(CONTENTS_ID[i]) + "_" + str(Lid) + ".mp4"
            print(fileName)
            # time.sleep(random.randrange(1, 60))

            # print(TARGET_DIR + fileName)
            ERROR_CNT = download(reqUrl, fileName, ERROR_CNT)
            if ERROR_CNT > 2:
                ERROR_CNT = 0
                break


