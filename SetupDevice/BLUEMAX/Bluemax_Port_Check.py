import openpyxl
import os, sys
import socket

from datetime import date
import time

CUR_PATH = os.getcwd()
SRC_FILE = CUR_PATH + "\\fwList.xlsx"
START_DATE = date.today()


def port_check(ip, port):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        sock.settimeout(1)
        result = sock.connect_ex((ip, port))
        sock.close()
    except socket.error as e:
        result = 1
        print("Error : " + e)

    if result == 0:
        return True
    else:
        return False


def write_log(string):
    dat = str(START_DATE).replace("-", "")
    f = open(CUR_PATH + f'\\log_{dat}.txt', "a+")
    now = time.localtime()
    cur_time = "%02d/%02d,%02d:%02d:%02d" % (now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)
    f.write(f'{cur_time} ::: {string}\n')
    f.close()


def printProgress(iteration, total, prefix='', suffix='', decimals=1, barLength=100):
    formatStr = "{0:." + str(decimals) + "f}"
    percent = formatStr.format(100 * (iteration / float(total)))
    filledLength = int(round(barLength * iteration / float(total)))
    bar = '#' * filledLength + '-' * (barLength - filledLength)
    sys.stdout.write('\r%s |%s| %s%s %s' % (prefix, bar, percent, '%', suffix)),
    if iteration == total:
        sys.stdout.write('\n')
    sys.stdout.flush()


if __name__ == "__main__":

    wb = openpyxl.load_workbook(SRC_FILE)
    ws = wb.active
    cnt = 0

    for row in range(2, ws.max_row + 1):
        schNo = ws[f'A{row}'].value
        schHaCls = ws[f'B{row}'].value
        schName = ws[f'D{row}'].value
        schUtmIp = ws[f'E{row}'].value
        schKTL3Ip = ws[f'F{row}'].value
        schKT_UTM_sIp = ws[f'G{row}'].value
        schKT_UTM_mIp = ws[f'H{row}'].value

        if schUtmIp and schKTL3Ip and schKT_UTM_sIp and schKT_UTM_mIp:
            skFwIp = schUtmIp
            ktFw1Ip = schKT_UTM_sIp
            ktFw2Ip = schKT_UTM_mIp
            ktL3Ip = schKTL3Ip
        else:
            write_log(f"{schNo}_{schName} / Ip address Empty")
            ws[f'I{row}'].value = f"Ip address Empty"
            continue

        kt_fw1_ck1 = port_check(ktFw1Ip, 22)
        kt_fw1_ck2 = port_check(ktFw1Ip, 443)
        kt_fw2_ck1 = port_check(ktFw2Ip, 22)
        kt_fw2_ck2 = port_check(ktFw2Ip, 443)

        if kt_fw1_ck1 and kt_fw1_ck2:
            ws[f'K{row}'].value = "Master OK"
        else:
            # write_log(f"{schNo}_{schName} / KT Slave 22,443 Port Closed")
            ws[f'K{row}'].value = f"Master Port Closed"

        if kt_fw2_ck1 and kt_fw2_ck2:
            ws[f'J{row}'].value = "Slave Ok"
        else:
            # write_log(f"{schNo}_{schName} / SK 22 Port Closed")
            ws[f'J{row}'].value = f"Slave Port Closed"

        # if port_check(ktFw1Ip, 22) and port_check(ktFw1Ip, 443) and port_check(ktFw2Ip, 22) and port_check(ktFw2Ip, 443):
        #     continue

        printProgress(row, ws.max_row, 'Progress:', 'Complete ', 1, 50)

        cnt += 1
        if cnt % 5 == 0:
            wb.save(SRC_FILE)

    wb.save(SRC_FILE)

