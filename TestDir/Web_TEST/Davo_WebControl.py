# parser.py
import requests

# Session 생성

MAIN_URL = 'http://192.168.0.226:88/login'
SETTING_URL = 'http://192.168.0.226:88/setting/wl5_wlan'
RESTART_URL = 'http://192.168.0.226:88/system/restart'

login_data = {
    'user_id': 'admin',
    'user_pwd': 'tmakxmdpdj~!@'
}

setting_data = {
    'disabled': '0',
    'ssid': 'SenWiFi_Free',
    'hidden': '1',
    'key': 'sen2021!wi',
    'key1': '',
    'key_type': 'ascii',
    'wep_key_type': 'hex',
    'wep_key_len': '64',
    'wpa_type': 'psk-mixed',
    'rsn_pairwise': 'aes',
    'vap_[]': 'vap00',
    'mode': 'none',
    'sta_auth': '0',
    'max_sta': '512',
    'encryption': 'psk-mixed+aes',
    'wmm': '1',
    'dummyVal': '2021723144302290',
    'act': 'set_info',
}

restart_data = {
    'dummyVal': '2021723144302290',
    'act': 'wifi_restart',
}

with requests.Session() as s:
    with s.post(MAIN_URL, data=login_data) as res:
        cookies = res.cookies.get_dict()
        print(cookies['ssid'])
        print(s.headers)
        print(res.text)

    with s.post(SETTING_URL, data=setting_data) as res:
        print(res.text)

    with s.post(RESTART_URL, data=restart_data) as res:
        print(res.text)
