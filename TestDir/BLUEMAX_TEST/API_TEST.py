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

CUR_PATH = os.getcwd()
SRC_FILE = CUR_PATH + "\\PolChkForm.xlsx"
START_DATE = date.today()

proxy = {'https': 'http://127.0.0.1:8080'}

# ### INTERFACE GET API
INTERFACE_API = '/api/sm/interfaces'

# ### VIP GET API
VIPS_API = '/api/sm/ha/virtual-ips'

# ### DHCP AREA GET API
DHCP_AREA_API = '/api/sm/dhcp/server/dynamic-areas'


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
    mainUrl = "https://192.168.10.10:50002"
    ha_api = f'/ha:50002:ha_grp_3'

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

        with s.get(mainUrl + INTERFACE_API, headers=headers, verify=False, proxies=proxy) as res:
            res_dict = json.loads(res.text)
            if_list = res_dict['result']
            for if_val in if_list:
                if if_val['ipv4address']:
                    for ip in if_val['ipv4address']:
                        print(if_val['name'])
                        print(ip)

        with s.get(mainUrl + ha_api + INTERFACE_API, headers=headers, verify=False, proxies=proxy) as res:
            res_dict = json.loads(res.text)
            if_list = res_dict['result']
            for if_val in if_list:
                if if_val['ipv4address']:
                    for ip in if_val['ipv4address']:
                        print(if_val['name'])
                        print(ip)

        with s.get(mainUrl + VIPS_API, headers=headers, verify=False, proxies=proxy) as res:
            res_dict = json.loads(res.text)
            vip_list = res_dict['result']
            for vip_val in vip_list:
                # print(vip_val)
                for val in vip_val['children']:
                    print(val['ifc_name'])
                    print(val['mmbr_uuid'])
                    print(val['ip'])
                    print(val['netmask'])

        with s.get(mainUrl + DHCP_AREA_API, headers=headers, verify=False, proxies=proxy) as res:
            res_dict = json.loads(res.text)
            dhcp_list = res_dict['result']
            # print(dhcp_list)
            for dhcp_val in dhcp_list:
                print(dhcp_val['net_addr'])
                print(dhcp_val['subnet_mask'])




