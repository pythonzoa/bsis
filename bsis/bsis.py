import requests
from io import BytesIO
from zipfile import ZipFile
from xml.etree.ElementTree import  parse
import pandas as pd
from lxml import html
from urllib.request import Request, urlopen
from urllib.parse import urlencode, quote_plus, unquote
from bs4 import BeautifulSoup

KEY = '825afc8affaa9a62a0a8e3425a6ce9cd71ab9c95'
url = 'https://opendart.fss.or.kr/api/corpCode.xml'
params = {'crtfc_key' : KEY}

response_list = requests.get(url, params=params).content

with ZipFile(BytesIO(response_list)) as zipfile:
    zipfile.extractall(('corpCode'))

# XML 호출하여 읽어오기
xmlTree = parse(r'corpCode\corpCode.xml')
root = xmlTree.getroot()
raw_list = root.findall('list')

corp_list = []

for i in range(0, len(raw_list)):
    corp_code = raw_list[i].findtext('corp_code')
    corp_name = raw_list[i].findtext('corp_name')
    stock_code = raw_list[i].findtext('stock_code')
    modify_date = raw_list[i].findtext('modify_date')

    corp_list.append([corp_code, corp_name, stock_code, modify_date])

# 정리
corp_list_df = pd.DataFrame(corp_list, columns=['Corp Code', 'Corp Name', 'Stock Code', 'Modify Date'])
print(corp_list_df.head())

corp_df = pd.DataFrame(corp_list, columns=['고유번호','정식명칭','종목코드','최종변경일자'])
print(corp_df.head())

stock_df = corp_df[corp_df['종목코드'] != ' ']
stock_df = stock_df[['고유번호','정식명칭','종목코드']].drop_duplicates()
print(stock_df.head())

def get_items(KEY, CORP_CODE, YEAR, RPT_CODE):
    url = 'https://opendart.fss.or.kr/api/fnlttSinglAcnt.xml'
    params = {'crtfc_key' : KEY,
              'corp_code' : CORP_CODE,
              'bsns_year' : YEAR,
              'reprt_code' : RPT_CODE}

    response = requests.get(url, params=params).content.decode('UTF-8')

    xml_obj = BeautifulSoup(response, 'html.parser')
    rows = xml_obj.findAll('list')
    return rows

item_list = [
'bsns_year', #	사업 연도	2019
'stock_code',#	종목 코드	상장회사의 종목코드(6자리)
'reprt_code',#	보고서 코드	1분기보고서 : 11013 반기보고서 : 11012 3분기보고서 : 11014 사업보고서 : 11011
'fs_div',#	개별/연결구분	OFS:재무제표, CFS:연결재무제표
'sj_div',#	재무제표구분	BS:재무상태표, IS:손익계산서
'account_nm',#  계정명	ex) 자본총계
'thstrm_nm',#	당기명	ex) 제 13 기 3분기말
'thstrm_dt',#	당기일자	ex) 2018.09.30 현재
'thstrm_amount',#	당기금액	9,999,999,999
'thstrm_add_amount',#	당기누적금액	9,999,999,999
'frmtrm_nm',#	전기명	ex) 제 12 기말
'frmtrm_dt',#	전기일자	ex) 2017.01.01 ~ 2017.12.31
'frmtrm_amount',#	전기금액	9,999,999,999
'frmtrm_add_amount',#	전기누적금액	9,999,999,999
'bfefrmtrm_nm',#	전전기명	ex) 제 11 기말(※ 사업보고서의 경우에만 출력)
'bfefrmtrm_dt',#	전전기일자	ex) 2016.12.31 현재(※ 사업보고서의 경우에만 출력)
'bfefrmtrm_amount',#	전전기금액	9,999,999,999(※ 사업보고서의 경우에만 출력)
'currency'#	통화 단위	통화 단위
]

YEAR = '2021'
RPT_CODE = '11011'
CORP_CODE = str('003490').zfill(8)
items = get_items(KEY,CORP_CODE,YEAR,RPT_CODE)

for i in range(0, len(items)):
    fs_item_list = []
    for item in item_list:
        try:
            value = items[i].find(item).text
        except:
            value = ''
        fs_item_list.append(value)
    print(CORP_CODE)
    
