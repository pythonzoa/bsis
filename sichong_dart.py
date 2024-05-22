import requests
from zipfile import ZipFile
from io import BytesIO
from xml.etree.ElementTree import parse
import pandas as pd
from bs4 import BeautifulSoup
import FinanceDataReader as fdr

# 상수
KEY = '825afc8affaa9a62a0a8e3425a6ce9cd71ab9c95'
URL_CORP_CODE = 'https://opendart.fss.or.kr/api/corpCode.xml'
URL_FIN_STATEMENT = 'https://opendart.fss.or.kr/api/fnlttSinglAcnt.xml'

def fetch_corp_codes(key):
    """회사 코드 데이터를 받아오고 파싱하는 함수"""
    response = requests.get(URL_CORP_CODE, params={'crtfc_key': key})
    with ZipFile(BytesIO(response.content)) as zipfile:
        zipfile.extractall('corpCode')
    xml_tree = parse('corpCode/CORPCODE.xml')
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
    corp_root = fetch_corp_codes(KEY)
    corp_df = prepare_corp_df(corp_root)
    quarter = "사업보고서"
    year = ['2016', '2017', '2018', '2019', '2020', '2021', '2022', '2023']
    bsis = "손익계산서"
    
    # 시가총액 1조원 이상의 종목 리스트 가져오기
    krx_list = fdr.StockListing('KRX')
    stock_cap = krx_list[krx_list['Marcap'] >= 1e12][['Code', 'Name']]
    
    for airline in stock_cap['Name']:
        if airline in corp_df['정식명칭'].values:
            corp_code, rpt_code, bsis_code = code_translator(corp_df, airline, quarter, bsis)
            
            fs_item_list = []
            for yr in year:
                financial_items = get_items(KEY, corp_code, yr, rpt_code)
                if financial_items:
                    for item in financial_items:
                        single_item_list = []
                        for tag_name in ['bsns_year', 'stock_code', 'reprt_code', 'fs_div', 'sj_div', 'account_nm', 'thstrm_nm',
                                        'thstrm_dt', 'thstrm_amount', 'thstrm_add_amount', 'frmtrm_nm', 'frmtrm_dt',
                                        'frmtrm_amount', 'frmtrm_add_amount', 'bfefrmtrm_nm', 'bfefrmtrm_dt', 'bfefrmtrm_amount',
                                        'currency']:
                            value = item.find(tag_name).text if item.find(tag_name) else ''
                            single_item_list.append(value)
                        fs_item_list.append(single_item_list)
                else:
                    print(f"데이터가 없습니다: {yr}, {airline}")

            # DataFrame 생성
            fs_df = pd.DataFrame(fs_item_list, columns=[
                '사업 연도', '종목 코드', '보고서 코드', '개별/연결구분', '재무제표구분', '계정명',
                '당기명', '당기일자', '당기금액', '당기누적금액', '전기명', '전기일자', '전기금액', '전기누적금액',
                '전전기명', '전전기일자', '전전기금액', '통화 단위'])

            # 콤마 제거
            amount_columns = ['당기금액', '당기누적금액', '전기금액', '전기누적금액', '전전기금액']
            for col in amount_columns:
                fs_df[col] = fs_df[col].str.replace(',', '')

            df_result = fs_df[(fs_df['개별/연결구분'] == 'OFS') & (fs_df['재무제표구분'] == bsis_code)]

            df_final = pd.DataFrame(df_result, columns=[
                '사업 연도', '계정명', '당기명', '당기일자', '당기금액', '당기누적금액',
                '전기명', '전기일자', '전기금액', '전기누적금액', '전전기명', '전전기일자', '전전기금액', '통화 단위'])
            print(df_final)
            # file_name = f'financial_statements_{airline}.xlsx'
            # df_final.to_excel(file_name, index=False)
            # print(f"Excel 파일로 저장되었습니다: {file_name}")

# 메인 함수 실행
submit()
