import pandas as pd
import requests
from io import BytesIO
from zipfile import ZipFile
from xml.etree.ElementTree import parse
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk

# API 키와 요청 URL
KEY = '825afc8affaa9a62a0a8e3425a6ce9cd71ab9c95'
URL_CORP_CODE = 'https://opendart.fss.or.kr/api/corpCode.xml'
URL_FIN_STATEMENT = 'https://opendart.fss.or.kr/api/fnlttSinglAcntAll.xml'


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
        'reprt_code': rpt_code,
        'fs_div' : 'OFS'
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
    bsis_code = {'손익계산서': 'CIS', '재무상태표': 'BS'}[bsis]
    return corp_code, quarter_code, bsis_code


def submit():
    """제출 버튼에 연결된 함수"""
    airline, year, quarter, bsis = airline_cb.get(), year_cb.get(), quarter_cb.get(), bsis_cb.get()
    corp_code, rpt_code, bsis_code = code_translator(corp_df, airline, quarter, bsis)
    financial_items = get_items(KEY, corp_code, year, rpt_code)
    if financial_items:
        fs_item_list = []
        for item in financial_items:
            single_item_list = []
            for tag_name in ['bsns_year', 'stock_code', 'reprt_code', 'fs_div', 'sj_div',
                             'account_id','account_detail',
                             'account_nm','thstrm_nm','thstrm_dt', 'thstrm_amount', 'thstrm_add_amount', 'frmtrm_nm', 'frmtrm_dt',
                             'frmtrm_amount', 'frmtrm_add_amount', 'bfefrmtrm_nm', 'bfefrmtrm_dt', 'bfefrmtrm_amount',
                             'currency']:
                value = item.find(tag_name).text if item.find(tag_name) else ''
                single_item_list.append(value)
            fs_item_list.append(single_item_list)


        # DataFrame 생성
        fs_df = pd.DataFrame(fs_item_list, columns=[
            '사업 연도', '종목 코드', '보고서 코드', '개별/연결구분', '재무제표구분',
            '계정ID','계정상세',
            '계정명','당기명', '당기일자', '당기금액', '당기누적금액', '전기명', '전기일자', '전기금액', '전기누적금액',
            '전전기명', '전전기일자', '전전기금액', '통화 단위'])

        print(fs_df)
        # 콤마 제거
        amount_columns = ['당기금액', '당기누적금액', '전기금액', '전기누적금액', '전전기금액']
        for col in amount_columns:
            fs_df[col] = fs_df[col].str.replace(',', '')

        df_result = fs_df[fs_df['재무제표구분'] == bsis_code]

        df_final = pd.DataFrame(df_result, columns=['사업 연도',
                                                    #'계정ID','계정상세',
                                                    '계정명','당기명', '당기일자', '당기금액', '당기누적금액',
        '전기명', '전기일자', '전기금액', '전기누적금액','전전기명', '전전기일자', '전전기금액', '통화 단위'])
        df_final.to_excel(f'{airline}.{year}.{quarter}.{bsis}.xlsx', index=False)
    else:
        print("데이터가 없습니다.")
    root.destroy()


# GUI 설정 및 메인 실행 루프
root = tk.Tk()
root.geometry("300x300")
root.title("재무제표 추출기")

airline_label = ttk.Label(root, text="항공사 선택:")
airline_label.pack(pady=(10, 0))
airline_cb = ttk.Combobox(root, values=["대한항공", "아시아나항공", "제주항공", "진에어", "티웨이항공", "에어부산"])
airline_cb.pack(pady=5)
airline_cb.set("대한항공")

year_label = ttk.Label(root, text="연도 선택:")
year_label.pack(pady=(10, 0))
year_cb = ttk.Combobox(root, values=[str(y) for y in range(2016, 2025)])
year_cb.pack(pady=5)
year_cb.set("2023")

quarter_label = ttk.Label(root, text="분기 선택:")
quarter_label.pack(pady=(10, 0))
quarter_cb = ttk.Combobox(root, values=["1분기보고서", "반기보고서", "3분기보고서", "사업보고서"])
quarter_cb.pack(pady=5)
quarter_cb.set("1분기보고서")

bsis_label = ttk.Label(root, text="재무제표 선택:")
bsis_label.pack(pady=(10, 0))
bsis_cb = ttk.Combobox(root, values=["손익계산서", "재무상태표"])
bsis_cb.pack(pady=5)
bsis_cb.set("손익계산서")

corp_root = fetch_corp_codes(KEY)
corp_df = prepare_corp_df(corp_root)

submit_button = ttk.Button(root, text="제출", command=submit)
submit_button.pack(pady=(10, 20))
root.mainloop()
