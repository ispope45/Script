import shutil
import win32com.client
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

# excel 불러오기
wb = excel.Workbooks.Open('C:\\Users\\Jungle\\Desktop\\test11.xlsx')
ws = wb.ActiveSheet

# 파일 복사 For문
for i in range(2, 68):
    shutil.copy(ws.Cells(i, 1).Value, 'C:\\Users\\Jungle\\Desktop\\지빠귀\\' + ws.Cells(i, 2).Value + '_SKB준공서류.xlsx')

# #파일 이동
# for i in range(2, 68):
#   shutil.move(ws.Cells(i, 1).Value, ws.Cells(i, 2).Value)
