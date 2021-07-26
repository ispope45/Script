import requests
from bs4 import BeautifulSoup as bs

URL = 'http://192.168.0.226:88/login'

login_data = {
    'user_id': 'admin',
    'user_pwd': 'tmakxmdpdj~!@'
}
with requests.Session() as s:
    first_req = s.post(URL)
    html = first_req.text
    soup = bs(html, 'html.parser')

    login_req = s.post(URL, data=login_data)
    





