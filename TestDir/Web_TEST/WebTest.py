import requests
import time
import base64
import binascii
import hashlib

from Crypto.Cipher import AES
from Crypto import Random
from bs4 import BeautifulSoup as bs
import json
MAIN_URL = 'https://192.168.10.11/'
URL = 'https://192.168.10.11/api/au/login'
login = {'id': 'admin', 'pw': 'secui00@!'}

proxy = {
    'https': 'http://127.0.0.1:8080'
}


def _pad(s):
    bs = 16
    return s + (bs - len(s) % bs) * chr(bs - len(s) % bs)

cookies = {
    '_next_login': 'index',
    '_static': '88d4bcd01551a342f397603aace6b3dd74fcab1d77800188e0b0e8bf058bbdf3'
}

with requests.Session() as s:
    with s.get(MAIN_URL, proxies=proxy, verify=False) as res:
        val = res.text
        find_num = val.find("csrf_token")
        csrf_token = val[find_num+51:find_num+115]

        iv = Random.new().read(AES.block_size)
        key = bytes(csrf_token[:32], 'utf-8')
        raw = bytes(_pad(login['pw']), 'utf-8')
        cipher = AES.new(key, AES.MODE_CBC, iv)
        hex_data = (iv.hex() + cipher.encrypt(raw).hex()).encode('utf-8')
        login_pw = base64.b64encode(hex_data).decode('utf-8')

        login_data = {
            "lang": "ko",
            "force": "true",
            "readonly": 0,
            "login_id": "admin",
            "login_pw": login_pw,
            "csrf_token": csrf_token,
            "type": "local"
        }

    with s.post(URL, json=login_data, cookies=cookies, verify=False, proxies=proxy) as res:
        print(res.cookies.get_dict())
        cookies = res.cookies.get_dict()
        print(res.headers)
        # print(type(cookies))
        # "_static = 88d4bcd01551a342f397603aace6b3dd74fcab1d77800188e0b0e8bf058bbdf3"
        print(cookies)
        val = res.json()
        # e_id = val['token']['id']
        print(res.text)
        # headers = {'X-Auth-Token': e_id, 'X-Requested-With': 'XMLHttpRequest', 'Accept-Encoding': 'gzip, deflate'}




