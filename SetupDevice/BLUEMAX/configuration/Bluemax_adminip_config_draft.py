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

# ### POLICY BATCH API
POLICY_BATCH_APPLY_API = '/api/po/fw/4/rules/batch'
pol_import_batch = {"file": "2ffbe3f5-1db5-48a9-a1b8-bf8ebf1b8644.xlsx",
                    "excel_pos": 1,
                    "allow_dup": 0,
                    "pre_rule_id": -1,
                    "top_rule_id_for_group": {"default": 5}}

POLICY_APPLY_API = '/api/po/command/fw-4-policies/apply'
pol_apply_json = {'mod_rule': 1}

# ### SNMP API CONFIG
SNMP_SETTING_API = '/api/sm/server/snmp/users'
snmp_setting_json = {"svr_name": "NMS",
                     "usr_name": "schoolnet4",
                     "security": 2,
                     "auth_type": 0,
                     "msg_enc_type": 0,
                     "auth_pw": "U2VuMTY3MDEzOTZA",
                     "msg_enc_pw": "U2VuMTY3MDEzOTZA"}

SNMP_V3_API = '/api/sm/server/snmp/v3'
snmp_v3_json = {"use": 1,
                "snmp_usr_id": 1}

SNMP_ENABLE_API = '/api/sm/server/snmp'
snmp_enable_json = {"use": 1,
                    "community": ""}

SNMP_APPLY_API = '/api/sm/command/server-snmp/apply'

# ### SYSLOG API
SYSLOG_SETTING_API = '/api/sm/server/syslog'
syslog_setting_json = {"svr1_use": 1,
                       "svr1_addr": "192.168.10.254",
                       "svr1_port": 7514,
                       "svr1_func": 0,
                       "svr1_form": 0,
                       "svr1_desc": "",
                       "svr2_use": 0,
                       "svr2_addr": "",
                       "svr2_port": 514,
                       "svr2_func": 0,
                       "svr2_form": 0,
                       "svr2_desc": "",
                       "svr3_use": 0,
                       "svr3_addr": "",
                       "svr3_port": 514,
                       "svr3_func": 0,
                       "svr3_form": 0,
                       "svr3_desc": "",
                       "svr4_use": 0,
                       "svr4_addr": "",
                       "svr4_port": 514,
                       "svr4_func": 0,
                       "svr4_form": 0,
                       "svr4_desc": "",
                       "svr5_use": 0,
                       "svr5_addr": "",
                       "svr5_port": 514,
                       "svr5_func": 0,
                       "svr5_form": 0,
                       "svr5_desc": ""}

SYSLOG_APPLY_API = '/api/sm/command/server-syslog/apply'

LOG_SETTING_API = '/api/lr/monitor/setup/logs'
log_setting_json = {"items": [
    {"log_id": 130001, "mod_id": 1, "name": "traffic_session", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 100, "mod_id": 3, "name": "fw_traffic", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 101, "mod_id": 3, "name": "fw_rule_traffic", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5, "mod_id": 3, "name": "nat_session", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 104, "mod_id": 3, "name": "nat_traffic", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 8, "mod_id": 3, "name": "ad_group_mapping", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 6, "mod_id": 3, "name": "userauth", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 7, "mod_id": 3, "name": "fqdn_object_management", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 1001, "mod_id": 9, "name": "ha_event", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 1100, "mod_id": 9, "name": "ha_traffic", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 1101, "mod_id": 9, "name": "ha_status", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5003, "mod_id": 9, "name": "iap_interworking", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5007, "mod_id": 9, "name": "zombie_pc_solution", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5004, "mod_id": 9, "name": "blacklist_regist", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5006, "mod_id": 9, "name": "line_management", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5100, "mod_id": 9, "name": "system_resource", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5101, "mod_id": 9, "name": "daemon", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5102, "mod_id": 9, "name": "interface_traffic", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 3103, "mod_id": 9, "name": "oversubscription", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 106, "mod_id": 9, "name": "qos", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5107, "mod_id": 9, "name": "vlan_traffic", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 10001, "mod_id": 10, "name": "alert", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0},
    {"log_id": 5001, "mod_id": 11, "name": "audit", "local_use": 1, "syslog1_use": 1, "syslog2_use": 0, "syslog3_use": 0, "syslog4_use": 0, "syslog5_use": 0, "snmptrap1_use": 0, "snmptrap2_use": 0, "snmptrap3_use": 0, "auto_backup_use": 0, "smtp_use": 0}]}

LOG_SETTING_APPLY_API = '/api/lr/command/setup-logs/apply'

# ### ADMIN IP ACL
ADMIN_IP_ACL_API = '/api/sm/manager/sys-ips'
admin_ip_acl_json = {"ip_ver": 1,
                     "ip": "1.1.1.3",
                     "prefix": 32,
                     "desc": ""}

ADMIN_IP_ACL_APPLY_API = '/api/sm/manager/sys-ips/apply'


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

        with s.post(mainUrl + ADMIN_IP_ACL_API, json=admin_ip_acl_json, headers=headers, verify=False, proxies=proxy) as res:
            print(res.text)

        with s.put(mainUrl + ADMIN_IP_ACL_APPLY_API, headers=headers, verify=False, proxies=proxy) as res:
            print(res.text)
