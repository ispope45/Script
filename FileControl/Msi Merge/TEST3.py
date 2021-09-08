import shutil
import sys
import os


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


try:
    file4 = resource_path("HCSFP32.ocx")
    file5 = resource_path("HCSFP64.ocx")

    src_file1 = "C:\Program Files\Adobe Flash Player\HCSFP64.ocx"
    src_file2 = "C:\Program Files\Adobe Flash Player\HCSFP32.ocx"

    dst_file1 = "C:\Program Files\Adobe Flash Player\HCSFP64_old.ocx"
    dst_file2 = "C:\Program Files\Adobe Flash Player\HCSFP32_old.ocx"

    print(file4)
    print(file5)

    print("Change Adobe Flash Player Resource")

    if os.path.isfile(src_file1):
        os.rename(src_file1, dst_file1)
    else:
        print(src_file1 + " Not Found")

    if os.path.isfile(src_file2):
        os.rename(src_file2, dst_file2)
    else:
        print(src_file2 + " Not Found")

    shutil.copy(file4, src_file1)
    shutil.copy(file5, src_file2)
    print("Complete")

except Exception as e:
    print(e)

