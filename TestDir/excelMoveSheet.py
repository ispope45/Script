from win32com.client import Dispatch
import glob

PATH = "C:\\Users\\Jungly\\Desktop\\"
src_path = PATH + "src\\"
dst_path = PATH + "dst\\"

f_list = glob.glob(PATH + "src\\*")

dic1 = {}
key = []

for file in f_list:
    values = file.split('\\')
    key = values[5].split('_')
    print(key)

    dic1[int(key[0])] = [key[2], key[3], key[4].replace('.xlsx', ''), file]
    '''
    dic1 = [양식구분, 관할경찰서, 기관명]
    '''

print(dic1)

for i in range(1143, 1197):
#for i in [5, 6, 30, 36, 41, 44]:
    print("********************")
    print(dic1[i])
    src = dic1[i][3]
    dst = dst_path + dic1[i][1] + ".xlsx"
    print(src)
    print(dst)

    xl = Dispatch("Excel.Application")
    xl.Visible = False  # You can remove this line if you don't want the Excel application to be visible
    
    src_wb = xl.Workbooks.Open(Filename=src)
    dst_wb = xl.Workbooks.Open(Filename=dst)

    if dic1[i][0] == "양식2":
        for j in range(1, 4):
            print("양식2")
            print(dic1[i][2] + "-" + str(j))

            dst_ws = dst_wb.Worksheets(1)
            src_ws = src_wb.Worksheets(j)
            src_ws.Name = dic1[i][2] + "-" + str(j)
            src_ws.Copy(Before=dst_wb.Worksheets('End'))

    elif dic1[i][0] == "양식5":
        print("양식5")
        print(dic1[i][2] + "-4")

        dst_ws = dst_wb.Worksheets(1)
        src_ws = src_wb.Worksheets(1)
        src_ws.Name = dic1[i][2] + "-4"
        src_ws.Copy(Before=dst_wb.Worksheets('End'))

    dst_wb.Close(SaveChanges=True)
    src_wb.Close(SaveChanges=True)
    xl.Quit()


    ''' 
    excel close
    '''



'''
    xl = Dispatch("Excel.Application")
    xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible
    
    src_wb = xl.Workbooks.Open(Filename=src)

'''





'''
xl = Dispatch("Excel.Application")
xl.Visible = True  # You can remove this line if you don't want the Excel application to be visible

src_wb = xl.Workbooks.Open(Filename=src)
dst_wb = xl.Workbooks.Open(Filename=dst)

dst_ws = dst_wb.Worksheets(1)
#ws2.Name = 'TEST7'

src_ws = src_wb.Worksheets(1)
src_ws.Copy(Before=dst_wb.Worksheets('End'))

dst_wb.Close(SaveChanges=True)
src_wb.Close(SaveChanges=True)
xl.Quit()

'''