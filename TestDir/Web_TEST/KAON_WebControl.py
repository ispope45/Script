# parser.py
import requests
import json
import hashlib

# Session 생성

MAIN_URL = 'http://192.168.0.227:8080/api/auth/token'
SETTING_URL = 'http://192.168.0.227:8080/api/wireless'
RESTART_URL = 'http://192.168.0.227:8080/api/operation/reboot'

login_data = {'auth': {"username": "YWRtaW4=", "authHash": ""}}

proxy = {
    'http': 'http://127.0.0.1:8080'
}

set_data = {"in-body":{"band-steering":True,"5g":{"radio":{"id":0,"bandwidth":7,"channel":0,"wps":True,"sideband":-1,"dfs":False,"dfs-status":"IDLE","airtime-frame":False,"power":"100","nphy":"-1","rate":"54000000","frag-threshold":"2346","rts-threshold":"2347","dtim-interval":"1","beacon-interval":"100","beacon-rotation":False,"preamble-type":"1","max-assoc":"64"},"wlans":[{"id":0,"ssid_id":0,"status":True,"ssid":"KAON_TEST_PY2","enc-method":"WPA2 Mixed","enc-algorithm":"aes","enc-key":"test43214321","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":True,"isolate":False,"wmm":True,"wmf":True},{"id":0,"ssid_id":1,"status":False,"ssid":"KaonAP-Guest01-5G","enc-method":"WPA2 Mixed","enc-algorithm":"tkip+aes","enc-key":"sen2021!wi","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":False,"isolate":False,"wmm":True,"wmf":True},{"id":0,"ssid_id":2,"status":False,"ssid":"KaonAP-Guest02-5G","enc-method":"WPA2 Mixed","enc-algorithm":"tkip+aes","enc-key":"sen2021!wi","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":False,"isolate":False,"wmm":True,"wmf":True},{"id":0,"ssid_id":3,"status":False,"ssid":"KaonAP-Guest03-5G","enc-method":"WPA2 Mixed","enc-algorithm":"tkip+aes","enc-key":"sen2021!wi","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":False,"isolate":False,"wmm":True,"wmf":True},{"id":0,"ssid_id":4,"status":False,"ssid":"KaonAP-Guest04-5G","enc-method":"WPA2 Mixed","enc-algorithm":"tkip+aes","enc-key":"sen2021!wi","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":False,"isolate":False,"wmm":True,"wmf":True}]},"2g":{"radio":{"id":1,"bandwidth":1,"channel":0,"wps":True,"sideband":"-1","airtime-frame":False,"power":"100","nphy":"-1","rate":"54000000","frag-threshold":"2346","rts-threshold":"2347","dtim-interval":"1","beacon-interval":"100","beacon-rotation":False,"preamble-type":"1","max-assoc":"64"},"wlans":[{"id":1,"ssid_id":0,"status":False,"ssid":"KAON_TEST1","enc-method":"WPA2 Mixed","enc-algorithm":"aes","enc-key":"test12341234","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":False,"isolate":False,"wmm":True,"wmf":True},{"id":1,"ssid_id":1,"status":False,"ssid":"KaonAP-Guest01-2.4G","enc-method":"WPA2 Mixed","enc-algorithm":"tkip+aes","enc-key":"sen2021!wi","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":False,"isolate":False,"wmm":True,"wmf":True},{"id":1,"ssid_id":2,"status":False,"ssid":"KaonAP-Guest02-2.4G","enc-method":"WPA2 Mixed","enc-algorithm":"tkip+aes","enc-key":"sen2021!wi","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":False,"isolate":False,"wmm":True,"wmf":True},{"id":1,"ssid_id":3,"status":False,"ssid":"KaonAP-Guest03-2.4G","enc-method":"WPA2 Mixed","enc-algorithm":"tkip+aes","enc-key":"sen2021!wi","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":False,"isolate":False,"wmm":True,"wmf":True},{"id":1,"ssid_id":4,"status":False,"ssid":"KaonAP-Guest04-2.4G","enc-method":"WPA2 Mixed","enc-algorithm":"tkip+aes","enc-key":"sen2021!wi","r-server-ip":"0.0.0.0","r-server-port":"1812","r-server-key":"1234","hidden":False,"isolate":False,"wmm":True,"wmf":True}]}}}

with requests.Session() as s:
    with s.get(MAIN_URL) as res:
        cookies = res.cookies.get_dict()
        print(cookies)
        val = res.json()
        challenge = val['auth']['challenge']
        key = "admin_password_" + challenge
        enc = hashlib.md5()
        enc.update(key.encode('utf-8'))
        encKey = enc.hexdigest()
        login_data['auth']['authHash'] = encKey
        print(login_data)

    with s.post(MAIN_URL, json=login_data) as res:
        cookies = res.cookies.get_dict()
        print(cookies)
        val = res.json()
        e_id = val['token']['id']
        print(res.text)
        headers = {'X-Auth-Token': e_id, 'X-Requested-With': 'XMLHttpRequest', 'Accept-Encoding': 'gzip, deflate'}

    with s.put(SETTING_URL, json=set_data, headers=headers, cookies=cookies) as res:
        cookies = res.cookies.get_dict()
        print(cookies)
        print(res.text)

    with s.post(RESTART_URL) as res:
        print(res.text)
