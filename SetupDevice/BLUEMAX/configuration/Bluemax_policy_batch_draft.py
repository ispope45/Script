import requests
import base64
import openpyxl
import json
import os, sys
import random, string

from requests_toolbelt import MultipartEncoder
from Cryptodome.Cipher import AES
from Cryptodome import Random

from datetime import date
import time
import urllib3
import random

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

LOGIN_API = '/api/au/login'
HA_API = '/ha:443:ha_grp_2'
POLICY_API = '/api/po/fw/4/rules?key='
INTERFACE_API = '/api/sm/interfaces'
POLICY_IMPORT_API = '/api/co/file/import'


# PING_API = '/api/co/tools/ping'

CUR_PATH = os.getcwd()
SRC_FILE = CUR_PATH + "\\PolChkForm.xlsx"
START_DATE = date.today()

proxy = {'https': 'http://127.0.0.1:8080'}

POLICY_BATCH_APPLY_API = '/api/po/fw/4/rules/batch'
pol_import_batch = {"file": "2ffbe3f5-1db5-48a9-a1b8-bf8ebf1b8644.xlsx",
                    "excel_pos": 1,
                    "allow_dup": 0,
                    "pre_rule_id": -1,
                    "top_rule_id_for_group": {"default": 5}}
POLICY_APPLY_API = '/api/po/command/fw-4-policies/apply'
pol_apply_json = {'mod_rule': 1}


def make_login_data(login_pw, csrf_token):
    login_data = {
        "lang": "ko",
        "force": "true",
        "readonly": 0,
        "login_id": "admin",
        "login_pw": login_pw,
        "csrf_token": csrf_token,
        "type": "local"
    }
    return login_data


def _pad(s):
    bs = 16
    return s + (bs - len(s) % bs) * chr(bs - len(s) % bs)


if __name__ == "__main__":
    mainUrl = "https://192.168.10.10"
    upload_file = "./1.xlsx"
    upload_data = open(upload_file, 'rb')
    files = {"file_name": upload_data}

    with requests.Session() as s:
        # WEB GUI Initialize
        with s.get(mainUrl, verify=False, proxies=proxy) as res:
            val = res.text
            find_num = val.find("csrf_token")
            csrf_token = val[find_num + 51:find_num + 115]
            pw = "secui00@!"
            iv = Random.new().read(AES.block_size)

            key = bytes(csrf_token[:32], 'utf-8')
            raw = bytes(_pad(pw), 'utf-8')
            cipher = AES.new(key, AES.MODE_CBC, iv)
            hex_data = (iv.hex() + cipher.encrypt(raw).hex()).encode('utf-8')
            login_pw = base64.b64encode(hex_data).decode('utf-8')

            login_data = make_login_data(login_pw, csrf_token)

        # Login Phase
        with s.post(mainUrl + LOGIN_API, json=login_data, verify=False, proxies=proxy) as res:
            cookies = res.cookies.get_dict()
            res_dict = json.loads(res.text)
            auth_key = res_dict['result']['api_token']
            headers = {'Authorization': auth_key}

        fields = {
            'file': ('1.xlsx', upload_data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        }
        boundary = '----WebKitFormBoundary' \
                   + ''.join(random.sample(string.ascii_letters + string.digits, 16))
        m = MultipartEncoder(fields=fields, boundary=boundary)
        pol_import_headers = {'Authorization': auth_key,
                              'Content-Type': m.content_type}

        with s.post(mainUrl + POLICY_IMPORT_API, data=m, verify=False, headers=pol_import_headers, proxies=proxy) as res:
            res_dict = json.loads(res.text)
            print(res_dict)
            file_name = res_dict['result']['file']
            pol_import_batch['file'] = file_name

            print(pol_import_batch)

        with s.post(mainUrl + POLICY_BATCH_APPLY_API, json=pol_import_batch, verify=False, headers=headers, proxies=proxy) as res:
            print(res.text)

        with s.put(mainUrl + POLICY_APPLY_API, json=pol_apply_json, verify=False, headers=headers, proxies=proxy) as res:
            print(res.text)


