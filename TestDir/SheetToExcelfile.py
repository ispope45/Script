from win32com.client import Dispatch
import glob
'''
엑셀 특정 시트를 찾아("2.설치 사진첩(공통)") 각각 별도파일로 저장
'''
f_list = glob.glob("C:\\Users\\Jungle\\Desktop\\TEST3\\*")
aa = f_list[0].split('\\')

for f in f_list:
    path1 = f
    aa = f.split('\\')
   # path2 = 'C:\\Users\\Jungle\\Desktop\\TEST33\\' + aa[5]

    xl = Dispatch("Excel.Application")
    xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible

    wb1 = xl.Workbooks.Open(Filename=path1)
    wb2 = xl.Workbooks.Add()

    ws1 = wb1.Worksheets("2.설치 사진첩(공통)")
    ws1.Copy(Before=wb2.Worksheets(1))
    wb2.SaveAs('C:\\Users\\Jungle\\Desktop\\TEST33\\V' + aa[5])
    wb2.Close(SaveChanges=True)
    wb1.Close(SaveChanges=False)
    xl.Quit()
