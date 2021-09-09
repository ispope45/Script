from datetime import date
import os


def write_log(string):
    startDate = date.today()
    curPath = os.getcwd()
    curDate = str(startDate).replace("-", "")
    f = open(curPath + f'\\log_{curDate}.txt', "a+")
    f.write(f'{string}\n')
    f.close()


if __name__ == "__main__":
    write_log("loglogloglog")