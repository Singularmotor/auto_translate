#coding = utf-8
#running in python3.6

import xlrd
import xlwt
import json
import http.client
import random
import hashlib
from urllib import parse
from time import sleep
from xlutils.copy import copy


def translate_baidu(orginal_text, orginal_lang, goal_lang):
    appid = 'xxxxx'  # 你的appid（百度申请）
    secretKey = 'xxxxx'  # 你的密钥（百度申请）
    text_translated = []
    dict_respond = None
    httpClient = None
    myurl = '/api/trans/vip/translate'
    q = orginal_text
    fromLang = 'en'
    toLang = 'zh'
    salt = random.randint(32768, 65536)

    sign = appid + q + str(salt) + secretKey
    m1 = hashlib.md5()
    m1.update(bytes(sign, encoding='utf-8'))
    sign = m1.hexdigest()
    myurl = myurl + '?appid=' + appid + '&q=' + parse.quote(
        orginal_text) + '&from=' + orginal_lang + '&to=' + goal_lang + '&salt=' + str(salt) + '&sign=' + sign

    try:
        httpClient = http.client.HTTPConnection('api.fanyi.baidu.com')
        httpClient.request('GET', myurl)

        # response是HTTPResponse对象
        response = httpClient.getresponse()
        rr = response.read()
        json_str = rr.decode('unicode_escape')
        print(1)
        dict_respond = json.loads(json_str)
        print(dict_respond)
        for i in dict_respond['trans_result']:
            print(2)
            text_translated.append(i['dst'])
            print(3)
    except Exception as e:
        print('错误:' + str(e))
        text_translated = str(e)
    finally:
        if httpClient:
            httpClient.close()

    return text_translated

readbook = xlrd.open_workbook('translate_original.xls') #翻译原文档
sheet = readbook.sheet_by_index(1)

book2 = copy(readbook)  # 拷贝一份原来的excel

writesheet = book2.get_sheet(1)
for j in range(2, sheet.ncols):
    for i in range(2, sheet.nrows):
        row_list = str(sheet.cell(i, 1).value)
        dd = translate_baidu(row_list, sheet.cell(1, 1).value, sheet.cell(1, j).value)
        writesheet.write(i, j, dd)

        sleep(1)#控制文本请求数（百度免费标准版）

book2.save('translated.xls')

'''
支持语言和翻译代码
zh	中文
en	英语
yue	粤语
wyw	文言文
jp	日语
kor	韩语
fra	法语
spa	西班牙语
th	泰语
ara	阿拉伯语
ru	俄语
pt	葡萄牙语
de	德语
it	意大利语
el	希腊语
nl	荷兰语
pl	波兰语
bul	保加利亚语
est	爱沙尼亚语
dan	丹麦语
fin	芬兰语
cs	捷克语
rom	罗马尼亚语
slo	斯洛文尼亚语
swe	瑞典语
hu	匈牙利语
cht	繁体中文
vie	越南语
'''

