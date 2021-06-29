import win32com.client
import os

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Open('C:\\Users\\Jungle\\Desktop\\3.xlsx')
ws = wb.ActiveSheet
for i in range(2, 48):
    script = ('config t' + '\n'
    'hostname ' + ws.Cells(i, 7).Value + '\n'
    'vlan 40' + '\n'
    'interface range ge1-10' + '\n'
    'switchport mode access' + '\n'
    'switchport access vlan 40' + '\n'
    'exit' + '\n'             
    'int vlan40' + '\n'
    'ip address ' + ws.Cells(i, 5).Value + '/24\n'
    'exit' + '\n'
    'interface range ge1-8' + '\n'
    'poe power-mode high-power' + '\n'
    'poe auto-power-up enable' + '\n'
    'poe enable' + '\n'
    'exit' + '\n'
    'ip route 0.0.0.0/0 ' + ws.Cells(i, 6).Value + '\n'
    'ntp server ' + ws.Cells(i, 9).Value + '\n'
    'ntp enable' + '\n'
    'logging timanager ' + ws.Cells(i, 8).Value + '\n' 
    'logging timanager 10.0.100.122' + '\n'
    'snmp-server community public ro' + '\n'
    'snmp-server community himalaya_lsf rw' + '\n'
    'snmp-server apply' + '\n'
    'exit' + '\n'
    'write memory' + '\n')

    print(script)
    f = open("C:\\Users\\Jungle\\Desktop\\NIA\\" + str(int(ws.Cells(i, 1).Value)) + "_" + ws.Cells(i, 2).Value + "_" + ws.Cells(i, 3).Value + ".txt", 'w')
    f.write(script)
    f.close()

excel.Quit()