from urllib.parse import urlencode, quote_plus

import json
import requests
import wget
import os
import math
import sys

CUR_PATH = os.getcwd()
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
        quote_plus('inqryBgnDt'): '200001010000',
        quote_plus('inqryEndDt'): '202110010000',
        quote_plus('type'): 'json'
    })
    req = url + queryParams
    reply = requests.get(req).text
    print(reply)

    jData = json.loads(reply)
    imgUrl = jData['response']['body']['items'][0]['prdctImgLrge']

    return imgUrl


def bar_custom(current, total, width=80):
    width = 30
    avail_dots = width - 2
    shaded_dots = int(math.floor(float(current) / total * avail_dots))
    percent_bar = '[' + '■'*shaded_dots + ' '*(avail_dots-shaded_dots) + ']'
    progress = "%d%% %s [%d / %d]" % (current / total * 100, percent_bar, current, total)
    return progress


def download(url, out_path):
    err = True
    try:
        wget.download(url, out=out_path, bar=bar_custom)
        print(" Ok")
        err = False
    except Exception as e:
        print(url + " Error : 404 Not Exist")
        print(e)
        f = open(DST_PATH + "\\wget_log.txt", "a+")
        f.write(out_path + f" Error : 404 Not Exist : {e}\n")
        f.close()
        err = True
    finally:
        return err


if __name__ == "__main__":
    imgUrl = g2b_prd_call("20922012")
    print(imgUrl)
    # if not os.path.isdir(resource_path("/img")):
    #     os.mkdir(resource_path("/img"))
    #
    # ERROR_CNT = download(imgUrl, resource_path("/img/20660748.jpg"))



