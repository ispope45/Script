import shutil
import sys
import os


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


try:
    # file1 = resource_path("1.msu")
    # file2 = resource_path("2.msi")
    # file3 = resource_path("3.msi")
    file4 = resource_path("HCSFP32.ocx")
    file5 = resource_path("HCSFP64.ocx")

    # file1 = "1.msu"
    # file2 = "2.msi"
    # file3 = "3.msi"

    # print(file1)
    print(file4)
    print(file5)

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
    print("Change Adobe Flash Player Resource")
    os.rename("C:\\Program Files\\Adobe Flash Player\\HCSFP64.ocx",
              "C:\\Program Files\\Adobe Flash Player\\HCSFP64_old.ocx")

    os.rename("C:\\Program Files\\Adobe Flash Player\\HCSFP32.ocx",
              "C:\\Program Files\\Adobe Flash Player\\HCSFP32_old.ocx")

    shutil.copy(file4, "C:\\Program Files\\Adobe Flash Player\\HCSFP32.ocx")
    shutil.copy(file5, "C:\\Program Files\\Adobe Flash Player\\HCSFP64.ocx")
    print("Complete")

except Exception as e:
    print(e)