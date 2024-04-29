import requests
from io import BytesIO
from zipfile import ZipFile
from xml.etree.ElementTree import parse
import pandas as pd
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk

airline = None
YEAR = None
quarter = None
bsis = None

def submit():
    global airline, YEAR, quarter, bsis
    airline = airline_cb.get()
    YEAR = year_cb.get()
    quarter = quarter_cb.get()
    bsis = bsis_cb.get()
    print(airline, YEAR, quarter, bsis)
    root.destroy()


root = tk.Tk()
root.geometry("300x300")
root.title("Hwi's 재무제표 추출")

# 항공사 선택 Combobox
airline_label = ttk.Label(root, text="항공사 선택:")
airline_label.pack(pady=(10, 0))

airline_cb = ttk.Combobox(root, values=["대한항공","아시아나항공","제주항공","진에어","티웨이항공","에어부산"])
airline_cb.pack(pady=5)
airline_cb.set("대한항공")  # 기본 선택값 설정

# 연도 선택 Combobox
year_label = ttk.Label(root, text="연도 선택:")
year_label.pack(pady=(10, 0))

year_cb = ttk.Combobox(root, values=['2016','2017','2018','2019','2020','2021','2022','2023','2024'])
year_cb.pack(pady=5)
year_cb.set("2023")  # 기본 선택값 설정

# 분기 선택 Combobox
quarter_label = ttk.Label(root, text="분기 선택:")
quarter_label.pack(pady=(10, 0))

quarter_cb = ttk.Combobox(root, values=["1분기보고서","반기보고서","3분기보고서","사업보고서"])
quarter_cb.pack(pady=5)
quarter_cb.set("1분기보고서")  # 기본 선택값 설정

# BSIS Combobox
bsis_label = ttk.Label(root, text="재무제표 선택:")
bsis_label.pack(pady=(10, 0))

bsis_cb = ttk.Combobox(root, values=["손익계산서", "재무상태표"])
bsis_cb.pack(pady=5)
bsis_cb.set("손익계산서")  # 기본 선택값 설정

# 제출 버튼
submit_button = ttk.Button(root, text="제출", command=submit)
submit_button.pack(pady=(10, 20))

root.mainloop()

################################################

# airline = '00113526'
# # 요청 변수 설정
# RPT_CODE = '11011'
# CORP_CODE = str(airline).zfill(8)
# bsis = 'BS'

def code_translator(airline, quarter, bsis):
    corp_df = corp_code  # assuming corp_df should be corp_code
    airline_code = corp_code['고유번호'][corp_df['정식명칭'] == airline]
    CORP_CODE = str(airline_code).zfill(8)
    quarter_code = {
       '1분기보고서' : '11013',
        '반기보고서' : '11012',
        '3분기보고서': '11014',
        '사업보고서' : '11011'
    }
    RPT_CODE = quarter_code[quarter]
    bsis_code = {
        '손익계산서' : 'IS',
        '재무상태표' : 'BS'
    }
    bsis = bsis_code[bsis]
    return CORP_CODE, RPT_CODE, bsis

# API 키와 요청 URL 설정
KEY = '825afc8affaa9a62a0a8e3425a6ce9cd71ab9c95'
url = 'https://opendart.fss.or.kr/api/corpCode.xml'
params = {'crtfc_key': KEY}

# API에서 ZIP 파일 받아오기
response_list = requests.get(url, params=params).content

# ZIP 파일 열고 XML 파일 추출하기
with ZipFile(BytesIO(response_list)) as zipfile:
    zipfile.extractall('corpCode')

# XML 파일 파싱
xmlTree = parse('corpCode/corpCode.xml')
root = xmlTree.getroot()
raw_list = root.findall('list')

# 회사 정보 리스트 생성
corp_list = []
for element in raw_list:
    corp_code = element.findtext('corp_code')
    corp_name = element.findtext('corp_name')
    stock_code = element.findtext('stock_code')
    modify_date = element.findtext('modify_date')
    corp_list.append([corp_code, corp_name, stock_code, modify_date])

# 데이터 프레임으로 변환 및 출력
corp_df = pd.DataFrame(corp_list, columns=['고유번호', '정식명칭', '종목코드', '최종변경일자'])

# 종목코드가 있는 데이터만 필터링
stock_df = corp_df[corp_df['종목코드'] != ' ']
stock_df = stock_df[['고유번호', '정식명칭', '종목코드']].drop_duplicates()
print(stock_df.head())


# 회계 정보를 가져오는 함수 정의
def get_items(KEY, CORP_CODE, YEAR, RPT_CODE):
    url = 'https://opendart.fss.or.kr/api/fnlttSinglAcnt.xml'
    params = {'crtfc_key': KEY,
              'corp_code': CORP_CODE,
              'bsns_year': YEAR,
              'reprt_code': RPT_CODE}
    response = requests.get(url, params=params).content.decode('UTF-8')
    xml_obj = BeautifulSoup(response, 'lxml-xml')
    return xml_obj.findAll('list')


code_translator(airline,quarter,bsis)
items = get_items(KEY, CORP_CODE, YEAR, RPT_CODE)
fs_item_list = [] # 회계 정보를 저장할 리스트 초기화

for item in items:
    single_item_list = [] # 개별 회계 정보를 저장할 임시 리스트
    for tag_name in ['bsns_year', 'stock_code', 'reprt_code', 'fs_div', 'sj_div', 'account_nm', 'thstrm_nm', 'thstrm_dt', 'thstrm_amount', 'thstrm_add_amount', 'frmtrm_nm', 'frmtrm_dt', 'frmtrm_amount', 'frmtrm_add_amount', 'bfefrmtrm_nm', 'bfefrmtrm_dt', 'bfefrmtrm_amount', 'currency']:
        value = item.find(tag_name).text if item.find(tag_name) else ''
        single_item_list.append(value)
        fs_item_list.append(single_item_list) # 각 회계 항목 정보를 리스트에 추가

# 모든 회계 정보가 수집된 후, pd.DataFrame을 사용하여 데이터 프레임을 생성합니다.
fs_df = pd.DataFrame(fs_item_list, columns=['bsns_year', 'stock_code', 'reprt_code', 'fs_div', 'sj_div', 'account_nm',
'thstrm_nm', 'thstrm_dt', 'thstrm_amount', 'thstrm_add_amount',
'frmtrm_nm', 'frmtrm_dt', 'frmtrm_amount', 'frmtrm_add_amount',
'bfefrmtrm_nm', 'bfefrmtrm_dt', 'bfefrmtrm_amount', 'currency'
])

print(fs_df.head()) # 결과 출력

fs_df = pd.DataFrame(fs_item_list, columns=[
    '사업 연도','종목 코드','보고서 코드','개별/연결구분','재무제표구분','계정명',
'당기명','당기일자','당기금액','당기누적금액','전기명','전기일자','전기금액','전기누적금액',
'전전기명','전전기일자','전전기금액','통화 단위'])
fs_df = fs_df.drop_duplicates()
print(fs_df.head())

df_result = fs_df[(fs_df['개별/연결구분']=='OFS')&(fs_df['재무제표구분']==bsis)]

df_final = pd.DataFrame(df_result, columns=[
    '사업 연도','보고서 코드','계정명',
'당기명','당기일자','당기금액','당기누적금액','전기명','전기일자','전기금액','전기누적금액',
'전전기명','전전기일자','전전기금액','통화 단위'])

df_final.to_excel(f'{airline}-{YEAR}.{RPT_CODE}.xlsx')