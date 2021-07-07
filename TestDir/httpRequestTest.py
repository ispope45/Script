import requests

URL = 'http://192.168.0.121:88/ap_login.asp'
# URL = 'http://192.168.0.121:88'

payload = {
    'http_passwd': 'tmakxmdpdj~!@',
}

# session = requests.session()
r = requests.post(URL, data=payload)
print(r.text)
