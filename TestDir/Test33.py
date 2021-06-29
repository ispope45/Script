from openpyxl import Workbook, load_workbook
import win32com.client
from win32com.client import Dispatch
import glob

PATH = "C:\\Users\\Jungly\\Desktop\\"
f_list = glob.glob(PATH + "src_ori\\*")

for file in f_list:
    xl = Dispatch("Excel.Application")
    xl.Visible = False  # You can remove this line if you don't want the Excel application to be visible
    #print(file)
    if file.find("~$") == -1:
        wb = xl.Workbooks.Open(Filename=file)
        ws = wb.Worksheets(1).Activate()
        sheetName = wb.Activesheet.Name
        #print(sheetName)

        if sheetName == "내보내기 요약":
            print("Success from " + file)

    wb.Close()
