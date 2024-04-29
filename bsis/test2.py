import requests
from io import BytesIO
from zipfile import ZipFile
from xml.etree.ElementTree import parse
import pandas as pd
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk

# API 키와 요청 URL
KEY = '825afc8affaa9a62a0a8e3425a6ce9cd71ab9c95'
URL_CORP_CODE = 'https://opendart.fss.or.kr/api/corpCode.xml'
URL_FIN_STATEMENT = 'https://opendart.fss.or.kr/api/fnlttSinglAcnt.xml'

def fetch_corp_codes(key):
    """회사 코드 데이터를 받아오고 파싱하는 함수"""
    response = requests.get(URL_CORP_CODE, params={'crtfc_key': key})
    with ZipFile(BytesIO(response.content)) as zipfile:
        zipfile.extractall('corpCode')
    xml_tree = parse('corpCode/corpCode.xml')
    return xml_tree.getroot()

def prepare_corp_df(root):
    """XML 데이터에서 회사 정보를 추출하여 DataFrame으로 변환"""
    corp_list = [
        (element.findtext('corp_code'), element.findtext('corp_name'),
         element.findtext('stock_code'), element.findtext('modify_date'))
        for element in root.findall('list')
    ]
    return pd.DataFrame(corp_list, columns=['고유번호', '정식명칭', '종목코드', '최종변경일자'])

def get_items(key, corp_code, year, rpt_code):
    """재무제표 데이터를 가져오는 함수"""
    params = {
        'crtfc_key': key,
        'corp_code': corp_code,
        'bsns_year': year,
        'reprt_code': rpt_code
    }
    response = requests.get(URL_FIN_STATEMENT, params=params)
    if response.status_code == 200:
        xml_obj = BeautifulSoup(response.content, 'lxml-xml')
        return xml_obj.findAll('list')
    else:
        return None

def code_translator(corp_df, airline, quarter, bsis):
    """사용자 입력을 코드로 변환"""
    airline_code = corp_df['고유번호'][corp_df['정식명칭'] == airline].values[0]
    corp_code = str(airline_code).zfill(8)
    quarter_code = {
        '1분기보고서': '11013',
        '반기보고서': '11012',
        '3분기보고서': '11014',
        '사업보고서': '11011'
    }[quarter]
    bsis_code = {'손익계산서': 'IS', '재무상태표': 'BS'}[bsis]
    return corp_code, quarter_code, bsis_code

def submit():
    """제출 버튼에 연결된 함수"""
    airline, year, quarter, bsis = airline_cb.get(), year_cb.get(), quarter_cb.get(), bsis_cb.get()
    corp_code, rpt_code, bsis = code_translator(airline,year, quarter, bsis)
    root.destroy()

# GUI 설정 및 메인 실행 루프
root = tk.Tk()
root.geometry("300x300")
root.title("Hwi's 재무제표 추출기")

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

submit_button = ttk.Button(root, text="제출", command=submit)
submit_button.pack(pady=(10, 20))
root.mainloop()

# 메인 실행 이후 데이터 처리
corp_root = fetch_corp_codes(KEY)
corp_df = prepare_corp_df(corp_root)
