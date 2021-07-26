# parser.py
import requests

# Session 생성

MAIN_URL = 'https://192.168.0.228:8080/admin/login/login'
SETTING_URL = 'https://192.168.0.228:8080/admin/wlan/wlan5gssid1'
# RESTART_URL = 'http://192.168.0.226:88/system/restart'

proxy = {
    'https': 'http://127.0.0.1:8080'
}

login_data = {
    'submit': '로그인',
    'loginid': 'admin',
    'loginpwd': 'allradio'
}

setting_data = {
    'wepkeyid': '',
    'enable': '1',
    'ssid': 'Allradio_TEST',  # SSID
    'hiddenssid': '1',
    'passphrase': 'test12341234',  # Password
    'viewpasswd': 'on',
    'wpakeyhex': '0',
    'authtype': 'wpapsk',
    'wepauthtype': 'auto',
    'wepkeylength': '5',
    'wepkeyhex': '1',
    'wepkey0': '',
    'dynamicwep': '0',
    'macauth': '0',
    'wpaversion': 'wpa1_2',
    'wpacipher': 'tkip%2Baes',
    'authsvraddr': '10.0.200.200',
    'authsvrport': '1812',
    'authsvrsecret': 'secret',
    'authsvraddr1': '',
    'authsvrport1': '',
    'authsvrsecret1': '',
    'vlanid': '',
    'assoccount': ''
}

with requests.Session() as s:
    with s.post(MAIN_URL, data=login_data, verify=False, proxies=proxy) as res:
        cookies = res.cookies.get_dict()
        print(cookies)
        print(s.headers)
        # print(res.text)

    with s.post(SETTING_URL, data=setting_data, verify=False, proxies=proxy) as res:
        print(res.text)

