import openpyxl
import os
import pandas as pd
import sys
from datetime import date

START_DATE = date.today()

SRC_DIR = os.getcwd() + "//"
DST_FILE = SRC_DIR + "result.xlsx"
COL_LIST1 = ['No', 'ORG', 'NAME']
COL_LIST2 = ['PRIORITY', 'ENABLED', 'SRC TYPE', 'SRC', 'SRC ADDR', 'DST TYPE', 'DST', 'DST ADDR', 'SVC', 'SVC SPEC',
             'ACTION', 'WEEKLY HIT COUNT', 'QUARTERLY HIT COUNT', 'REVERSIBLE']


def printProgress(iteration, total, prefix='', suffix='', decimals=1, barLength=100):
    formatStr = "{0:." + str(decimals) + "f}"
    percent = formatStr.format(100 * (iteration / float(total)))
    filledLength = int(round(barLength * iteration / float(total)))
    bar = '#' * filledLength + '-' * (barLength - filledLength)
    sys.stdout.write('\r%s |%s| %s%s %s' % (prefix, bar, percent, '%', suffix)),
    if iteration == total:
        sys.stdout.write('\n')
    sys.stdout.flush()


def write_log(string):
    dat = str(START_DATE).replace("-", "")
    f = open(SRC_DIR + f'log_{dat}.txt', "a+")
    f.write(f'{string}\n')
    f.close()


if __name__ == "__main__":

    fileList = os.listdir(SRC_DIR)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(COL_LIST1 + COL_LIST2)

    prog = 0
    for f in fileList:
        if f.find(".csv") == -1:
            continue
        fileNo = f.split("_")[0]
        fileOrg = f.split("_")[1]
        fileSch = f.split("_")[2].split(".")[0]

        data = pd.read_csv(SRC_DIR + f)

        p_data = data[COL_LIST2]

        p_data['No'] = fileNo
        p_data['ORG'] = fileOrg
        p_data['NAME'] = fileSch
        e_data = p_data[COL_LIST1 + COL_LIST2]

        totalValue = e_data.values.tolist()

        for i in totalValue:
            ws.append(i)
            # print(i)
        prog += 1
        printProgress(prog, len(fileList), 'Progress:', 'Complete ', 1, 50)

    wb.save(DST_FILE)
