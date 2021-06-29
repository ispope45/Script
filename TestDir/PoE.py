import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:\\Users\\Jungly\\Desktop\\4.xlsx')
ws = wb.ActiveSheet
for i in range(2, 180):
    script = (
    f'en\n'
    f'config terminal\n'
    f'hostname {str(ws.Cells(i, 4).Value)}\n'
    f'service ssh\n'
    f'ip ssh port 9992\n'
    f'interface vlan 1\n'
    f'ip address {str(ws.Cells(i, 3).Value)} {str(ws.Cells(i, 6).Value)}\n'
    f'exit\n'
    f'ip route 0.0.0.0/0 ' + str(ws.Cells(i, 5).Value) + '\n'
    f'end\n'
    f'write memory\n'
    f'y' + '\n\n'
    f'---------------- 추가설정(직접입력) ----------------' + '\n\n' 
    f'conf t' + '\n'
    f'snmp-server community RO' + '\n'
    f'aiedu' + '\n'
    f'aiedu' + '\n'
    f'snmp-server community RW' + '\n'
    f'niaedu' + '\n'
    f'niaedu' + '\n\n'
    f'write memory' + '\n'
    f'y' + '\n\n'
    f'---------------- 계정정보(직접입력) ----------------' + '\n\n'
    f'기본값 : root // [패스워드없음]' + '\n'
    f'** root 입력시 계정재생성 메세지출력됨' + '\n'
    f'변경값 : admin // frontier1!' + '\n'
    f' ')

    print(script)
    f = open("C:\\Users\\Jungly\\Desktop\\dst\\" + str(int(ws.Cells(i, 1).Value)) + "_" + str(ws.Cells(i, 2).Value) + "_PoE Config.txt", 'w+')
    f.write(script)
    f.close()

excel.Quit()