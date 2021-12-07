import requests
import openpyxl
import os
import sys

MAIN_URL = "https://192.168.0.254:50005/"
LOGIN_API = "login/loginAction"
BACKUP_API = "firewall/firewall/ipv4Policy/list?json=true&exportAction=true&queryString=&"

proxy = {'https': 'http://127.0.0.1:8080'}

login_data = {'disconnectType': '',
              'webDeviceForceLogin': '',
              'force': 'false',
              'userId': 'admin',
              'password': 'qwe123!@#',
              'language': 0}
CUR_PATH = os.getcwd()
SRC_DIR = CUR_PATH + "\\"
SRC_FILE = "form.xlsx"


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
    wb = openpyxl.load_workbook(SRC_DIR + SRC_FILE)
    ws = wb.active

    for row in range(2, ws.max_row + 1):
        no = str(ws[f'A{row}'].value)
        org = ws[f'B{row}'].value
        name = ws[f'C{row}'].value
        utmIp = ws[f'D{row}'].value

        MAIN_URL = f"https://{utmIp}:50005/"

        with requests.Session() as s:
            with s.post(MAIN_URL + LOGIN_API, data=login_data, verify=False) as res:
                print(res.status_code)

            with s.get(MAIN_URL + BACKUP_API, verify=False) as res:
                print(res.status_code)
                result = res.text.replace("\n\n", "\n")
                f = open(SRC_DIR + f"{no}_{org}_{name}.csv", 'w', newline='')
                f.write(result)
                f.close()
        s.close()

        printProgress(row, ws.max_row, 'Progress:', 'Complete ', 1, 50)
    wb.close()
    os.system("pause")
