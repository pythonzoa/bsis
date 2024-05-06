import pandas as pd
import requests
from io import BytesIO
from zipfile import ZipFile
from xml.etree.ElementTree import parse
import pyautogui
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import ttk, simpledialog, messagebox
import datetime

# API 키와 데이터 요청 URL
KEY = '825afc8affaa9a62a0a8e3425a6ce9cd71ab9c95'
URL_CORP_CODE = 'https://opendart.fss.or.kr/api/corpCode.xml'
URL_FIN_STATEMENT = 'https://opendart.fss.or.kr/api/fnlttSinglAcnt.xml'
URL_FIN_STATEMENT_D = 'https://opendart.fss.or.kr/api/fnlttSinglAcntAll.xml'

# 현재 연도를 가져오기
current_year = datetime.datetime.now().year

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

def get_items(key, corp_code, year, rpt_code, detail):
    """재무제표 데이터를 가져오는 함수"""
    params = {
        'crtfc_key': key,
        'corp_code': corp_code,
        'bsns_year': year,
        'reprt_code': rpt_code,
        'fs_div': 'OFS'
    }
    if detail == '단순':
        response = requests.get(URL_FIN_STATEMENT, params=params)
    else:
        response = requests.get(URL_FIN_STATEMENT_D, params=params)
    if response.status_code == 200:
        xml_obj = BeautifulSoup(response.content, 'lxml-xml')
        return xml_obj.findAll('list')
    else:
        return None

def code_translator(corp_df, airline, quarter, bsis, detail):
    """사용자 입력을 코드로 변환"""
    try:
        airline_code = corp_df['고유번호'][corp_df['정식명칭'] == airline].values[0]
    except IndexError:
        messagebox.showwarning("오류", f"입력한 회사명({airline})이(가) 없습니다. 다시 입력해주세요.")
        return None, None, None
    corp_code = str(airline_code).zfill(8)
    quarter_code = {
        '1분기보고서': '11013',
        '반기보고서': '11012',
        '3분기보고서': '11014',
        '사업보고서': '11011'
    }[quarter]
    if detail == '단순':
        bsis_code = {'손익계산서': 'IS', '재무상태표': 'BS'}[bsis]
    else:
        # 상세 옵션 선택 시, '손익계산서'의 경우 'CIS'를 사용하되, 없으면 'IS'를 사용
        if bsis == '손익계산서':
            try:
                bsis_code = {'손익계산서': 'CIS', '재무상태표': 'BS'}[bsis]
            except KeyError:
                bsis_code = 'IS'
        else:
            bsis_code = 'BS'
    return corp_code, quarter_code, bsis_code

def submit():
    """제출 버튼에 연결된 함수"""
    airline = airline_cb.get() if selected_option.get() == 'airline' else other_airline_name.get()
    year, quarter, bsis, detail = year_cb.get(), quarter_cb.get(), bsis_cb.get(), detail_cb.get()
    corp_root = fetch_corp_codes(KEY)
    corp_df = prepare_corp_df(corp_root)
    corp_code, rpt_code, bsis_code = code_translator(corp_df, airline, quarter, bsis, detail)
    if corp_code is None:
        return
    financial_items = get_items(KEY, corp_code, year, rpt_code, detail)
    if financial_items:
        fs_item_list = [
            [item.find(tag_name).text if item.find(tag_name) else '' for tag_name in [
                'bsns_year', 'stock_code', 'reprt_code', 'fs_div', 'sj_div', 'account_nm', 'thstrm_nm',
                'thstrm_dt', 'thstrm_amount', 'thstrm_add_amount', 'frmtrm_nm', 'frmtrm_dt',
                'frmtrm_amount', 'frmtrm_add_amount', 'bfefrmtrm_nm', 'bfefrmtrm_dt', 'bfefrmtrm_amount',
                'currency']]
            for item in financial_items]
        fs_df = pd.DataFrame(fs_item_list, columns=[
            '사업 연도', '종목 코드', '보고서 코드', '개별/연결구분', '재무제표구분', '계정명',
            '당기명', '당기일자', '당기금액', '당기누적금액', '전기명', '전기일자', '전기금액', '전기누적금액',
            '전전기명', '전전기일자', '전전기금액', '통화 단위'])

        if detail == '단순':
            df_result = fs_df[(fs_df['개별/연결구분'] == 'OFS') & (fs_df['재무제표구분'] == bsis_code)]
        else:
            df_result = fs_df[fs_df['재무제표구분'] == bsis_code]

        df_final = pd.DataFrame(df_result, columns=['사업 연도', '계정명', '당기명', '당기일자', '당기금액', '당기누적금액',
                                                    '전기명', '전기일자', '전기금액', '전기누적금액', '전전기명', '전전기일자', '전전기금액', '통화 단위'])
        df_final.to_excel(f'{airline}.{year}.{quarter}.{bsis}({detail}).xlsx', index=False)
    else:
        messagebox.showinfo("정보", "요청한 데이터가 없습니다.")
    root.destroy()

def show_option():
    """라디오 버튼 선택에 따라 옵션을 보여주는 함수"""
    if selected_option.get() == 'airline':
        airline_cb.config(state='readonly')  # 항공사 선택 가능
    else:
        airline_cb.config(state='disabled')  # 항공사 선택 비활성화
        request_company_name()

class CustomInputDialog(tk.Toplevel):
    """사용자 정의 입력 대화 상자 클래스"""
    def __init__(self, parent, title, prompt):
        super().__init__(parent)
        self.title(title)
        self.geometry("300x150")  # 대화 상자 크기 설정

        self.var = tk.StringVar(self)  # 입력 값을 저장할 변수

        ttk.Label(self, text=prompt).pack(padx=20, pady=10)  # 프롬프트 레이블

        # 입력 필드
        entry = ttk.Entry(self, textvariable=self.var)
        entry.pack(padx=20, pady=10, fill=tk.X)
        entry.focus()

        # 확인 버튼
        ttk.Button(self, text="확인", command=self.confirm).pack(pady=10)

        self.protocol("WM_DELETE_WINDOW", self.on_close)  # 사용자가 창을 닫을 때 이벤트 처리
        self.result = None

    def confirm(self):
        self.result = self.var.get()  # 입력받은 값을 저장
        self.destroy()  # 대화 상자 종료

    def on_close(self):
        self.result = None  # 창이 닫힐 때 결과를 None으로 설정
        self.destroy()  # 대화 상자 종료

def request_company_name():
    dialog = CustomInputDialog(root, "기타 회사", "회사 이름을 입력하세요:")
    pyautogui.press('hangul')
    root.wait_window(dialog)  # 대화 상자가 닫힐 때까지 대기
    pyautogui.press('hangul')

    company_name = dialog.result  # 입력받은 결과를 가져옴
    if company_name:
        other_airline_name.set(company_name)  # 결과가 있으면 설정
    else:
        # 입력이 취소되거나 값이 없는 경우
        selected_option.set('airline')  # 라디오 버튼을 '항공사'로 설정
        airline_cb.config(state='readonly')  # 항공사 콤보박스를 다시 활성화

# GUI 설정 및 메인 실행 루프
root = tk.Tk()
root.geometry("350x400")
root.title("재무제표 추출기")
selected_option = tk.StringVar(value='airline')
other_airline_name = tk.StringVar()

airline_radio = ttk.Radiobutton(root, text='항공사', variable=selected_option, value='airline', command=show_option)
airline_radio.pack(pady=(10, 0))

other_radio = ttk.Radiobutton(root, text='기타', variable=selected_option, value='other', command=show_option)
other_radio.pack(pady=5)

airline_cb = ttk.Combobox(root, values=["대한항공", "아시아나항공", "제주항공", "진에어", "티웨이항공", "에어부산"])
airline_cb.pack(pady=5)
airline_cb.set("대한항공")

year_label = ttk.Label(root, text="연도 선택:")
year_label.pack(pady=(10, 0))
year_cb = ttk.Combobox(root, values=[str(y) for y in range(2016, current_year + 1)])
year_cb.pack(pady=5)
year_cb.set(str(current_year))

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

detail_label = ttk.Label(root, text="상세 선택:")
detail_label.pack(pady=(10, 0))
detail_cb = ttk.Combobox(root, values=["단순", "상세"])
detail_cb.pack(pady=5)
detail_cb.set("단순")

submit_button = ttk.Button(root, text="추출", command=submit)
submit_button.pack(pady=(10, 20))

root.mainloop()
