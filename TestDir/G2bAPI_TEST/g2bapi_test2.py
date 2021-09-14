from urllib.parse import urlencode, quote_plus

import json
import requests
import wget
import os
import math
import sys
import openpyxl


CUR_PATH = os.getcwd() + '\\'
DST_PATH = CUR_PATH + '\\'


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def g2b_prd_call(prdNo):
    url = 'http://apis.data.go.kr/1230000/ThngListInfoService/getThngPrdnmLocplcAccotListInfoInfoPrdlstSearch'
    queryParams = '?' + urlencode({
        quote_plus('serviceKey'):
            'BE+Tsy32/R43W0b5Ab8rkDXpPkQyplKcrVghwj3OkFjNwiVp5J5FHXnhwm4eGBN583zA0/gRPpPpwsFa29n1Rg==',
        quote_plus('numOfRows'): '10',
        quote_plus('pageNo'): '1',
        quote_plus('prdctIdntNo'): prdNo,
        quote_plus('inqryBgnDt'): '199501010000',
        quote_plus('inqryEndDt'): '202110010000',
        quote_plus('type'): 'json'
    })
    req = url + queryParams
    reply = requests.get(req).text

    jData = json.loads(reply)
    if jData['response']['body']['totalCount'] > 0:
        data = jData['response']['body']['items'][0]['prdctImgLrge']
        res = True
    else:
        data = "Not in List"
        res = False

    return data, res


def download(url, out_path):
    try:
        wget.download(url, out=out_path)
    except Exception as e:
        f = open(DST_PATH + "\\img_log.txt", "a+")
        f.write(out_path + f"Error;404;{url};{e}\n")
        f.close()


if __name__ == "__main__":
    load_wb = openpyxl.load_workbook(CUR_PATH + 'raw.xlsx')
    load_ws = load_wb.active

    img_val = []
    img_prd = []
    for i in range(8, load_ws.max_row - 5):
        row = str(i)
        # print(load_ws['A' + row].value)
        prdCls = load_ws['A' + row].value
        img_val.append(prdCls)
        # tmp = prdCls.split('-')
        # print(tmp)
        img_prd.append(prdCls.split('-')[1])

    # print(img_prd)
    for val in img_prd:
        imgUrl, res = g2b_prd_call(val)

        if not os.path.isdir(resource_path("/img")):
            os.mkdir(resource_path("/img"))
        if not os.path.isfile(resource_path(f"/img/{val}.jpg")):
            download(imgUrl, resource_path(f"/img/{val}.jpg"))



