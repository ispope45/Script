from datetime import date
import os

START_DATE = date.today()
CUR_PATH = os.getcwd()

print(START_DATE)
print(type(START_DATE))
dat = str(START_DATE).replace("-", "")
print(dat)


def write_log(string):
    f = open(CUR_PATH + f'log_{dat}.txt', "a+")
    f.write(f'{string}\n')
    f.close()
