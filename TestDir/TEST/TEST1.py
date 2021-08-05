import os
import time

import sys
import os


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


try:
    # file1 = resource_path("1.msu")
    file2 = resource_path("2.msi")
    file3 = resource_path("3.msi")

    # file1 = "1.msu"
    # file2 = "2.msi"
    # file3 = "3.msi"

    # print(file1)
    print(file2)
    print(file3)

    # print("1/3 install...")
    # res1 = os.system("wusa /quiet " + file1)
    # # res1 = os.system("wusa " + file1)
    # chk = True
    # while chk:
    #     time.sleep(3)
    #     psChk = os.popen("TASKLIST | FINDSTR \"wusa\" | FINDSTR \"Console\"").read()
    #     psChk2 = psChk.find("wusa")
    #     if psChk2 == -1:
    #         chk = False

    print("1/2 install...")
    # res2 = os.system("msiexec /i " + file2)
    res2 = os.system("msiexec /i " + file2 + " /quiet")
    # os.execl(file2, '/i')

    chk = True
    while chk:
        time.sleep(3)
        psChk = os.popen("TASKLIST | FINDSTR \"msiexec\" | FINDSTR \"Console\"").read()
        psChk2 = psChk.find("msiexec")
        if psChk2 == -1:
            chk = False

    print("2/2 install...")
    # res3 = os.system("msiexec /i " + file3)
    res3 = os.system("msiexec /i " + file3 + " /quiet")

    while chk:
        time.sleep(3)
        psChk = os.popen("TASKLIST | FINDSTR \"msiexec\" | FINDSTR \"Console\"").read()
        psChk2 = psChk.find("msiexec")
        if psChk2 == -1:
            chk = False

    print("Complete")

except Exception as e:
    print(e)