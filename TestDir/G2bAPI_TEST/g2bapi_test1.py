from urllib.parse import urlencode, quote_plus
import requests

# API 승인대기
#
# url = 'http://apis.data.go.kr/1230000/ShoppingMallPrdctInfoService03/getUcntrctPrdctInfoList'
#
# queryParams = '?' + urlencode({
#     quote_plus('serviceKey'): 'BE+Tsy32/R43W0b5Ab8rkDXpPkQyplKcrVghwj3OkFjNwiVp5J5FHXnhwm4eGBN583zA0/gRPpPpwsFa29n1Rg==',
#     quote_plus('numOfRows'): '1',
#     quote_plus('pageNo'): '1',
#     quote_plus('prdctIdntNo'): '22748948',
#     quote_plus('prodctCertYn'): 'Y'
# })
#

url = 'http://apis.data.go.kr/1230000/ShoppingMallPrdctInfoService03/getShoppingMallPrdctInfoList'
queryParams = '?' + urlencode({
    quote_plus('serviceKey'): 'BE+Tsy32/R43W0b5Ab8rkDXpPkQyplKcrVghwj3OkFjNwiVp5J5FHXnhwm4eGBN583zA0/gRPpPpwsFa29n1Rg==',
    quote_plus('numOfRows'): '1',
    quote_plus('pageNo'): '1',
    quote_plus('inqryDiv'): '1',
    quote_plus('inqryBgnDate'): '20200901',
    quote_plus('inqryEndDate'): '20201001'
})

req = url + queryParams
print(req)

test = requests.get(req)

rep = test.text
print(rep)