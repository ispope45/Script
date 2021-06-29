from win32com.client import Dispatch
import glob
'''
엑셀 특정 시트를 찾아("2.설치 사진첩(공통)") 각각 별도파일로 저장
'''
f_list = glob.glob("C:\\Users\\Jungly\\Desktop\\독수리\\*")

for f in f_list:
    print(f)
    xl = Dispatch("Excel.Application")

    wb1 = xl.Workbooks.Open(Filename=str(f))
    ws1 = wb1.Worksheets("설치 사진첩")
    ws1 = wb1.ActiveSheet

    wb1.Close(SaveChanges=False)
    xl.Quit()
