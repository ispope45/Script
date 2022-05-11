import requests
import time
from bs4 import BeautifulSoup as bs

proxy = {
    'https': 'http://127.0.0.1:8080'
}

# ### Reservation
reservation_date = '2022-05-23'
reservation_time = [7, 8, 11, 12]
reservation_course = '1'  # 1 : LAKE, 2 : VALLEY, 3 : MOUNTAIN

MAIN_URL = "https://www.sejongemerson.co.kr"
LOGIN_API = "/member/member_loginOK.asp"
LOGIN_DATA = {
    'backaction': '/reservation/real_calendar.asp',
    'login_id': '501913',
    'pwd': '501913'
}

RESERVATION_LIST_API = "/reservation/real_calendar.asp"

RESERVATION_API = "/reservation/real_timeList.asp"
RESERVATION_DATA = {
    'txtFlag': 'getReservYN',
    'txtDate': '2022-05-10',
    'txtCourse': '',
    'txtCourseDesc': '',
    'txtTime': '',
    'txtRsvnNo': '',
    'txtMode': 'I'
}

RESERVATION_PROC_API = "/reservation/reservation_proc.asp"
RESERVATION_PROC_DATA = {
    'txtFlag': 'setGolfTime',
    'txtDate': '2022-05-18',
    'txtCourse': '1',
    'txtCourseDesc': '',
    'txtTime': '1753',
    'txtStat': 'RR',
    'txtRsvnNo': '',
    'txtMode': 'I',
    'txtCaddy': 'Y',
    'txtRemark': '',
    'tel1': '010',
    'tel2': '2775',
    'tel3': '4912'
}
# def date_parse(date):


def time_parse(time):
    time_val = time.replace(":", "")
    time_hour = time.split(":")[0]

    return time_val, int(time_hour)


if __name__ == "__main__":
    with requests.Session() as s:
        with s.post(MAIN_URL + LOGIN_API, data=LOGIN_DATA, verify=False) as res:
            reservation_vali = False

        while not reservation_vali:
            date_vali = False

            while not date_vali:
                time.sleep(0.5)

                with s.get(MAIN_URL + RESERVATION_LIST_API, verify=False) as res:
                    soup = bs(res.text, "html.parser")
                    elements = soup.select('div.mR-calendar a.mR-open')
                    for index, element in enumerate(elements, 1):
                        print(f"{element.attrs['date']}")
                        if element.attrs['date'] == reservation_date:
                            RESERVATION_DATA['txtDate'] = reservation_date
                            date_vali = True
                            break
                        else:
                            date_vali = False
                if date_vali:
                    break

            with s.post(MAIN_URL + RESERVATION_API, data=RESERVATION_DATA, verify=False) as res:
                soup = bs(res.text, "html.parser")
                elements = soup.select('table.mR-table-1 p.resv')
                # print(elements)

                for index, element in enumerate(elements, 1):
                    print(f"{element.attrs['date']}/{element.attrs['time']}/{element.attrs['course']}")
                    time_val, time_hour = time_parse(element.attrs['time'])
                    if time_hour in reservation_time and element.attrs['course'] == reservation_course:
                        RESERVATION_PROC_DATA['txtDate'] = reservation_date
                        RESERVATION_PROC_DATA['txtCourse'] = element.attrs['course']
                        RESERVATION_PROC_DATA['txtTime'] = time_val
                        reservation_vali = True
                        print(RESERVATION_PROC_DATA)
                        break

            if reservation_vali:
                print("reservation")
                # with s.post(MAIN_URL + RESERVATION_PROC_API, data=RESERVATION_PROC_DATA, verify=False) as res:
                #     print(res.text)
                #     if int(res.status_code) == 200:
                #         print("reservation OK")
                #         break
                #
                #     else:
                #         reservation_vali = False
