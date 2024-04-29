import tkinter as tk
from tkinter import ttk

airline = None
year = None
quarter = None
bsis = None
def submit():
    global airline, year, quarter, bsis
    airline = airline_cb.get()
    year = year_cb.get()
    quarter = quarter_cb.get()
    bsis = bsis_cb.get()
    print(airline, year, quarter, bsis)
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