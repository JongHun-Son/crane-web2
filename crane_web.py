# 📁 app.py
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="중량물 작업계획서", layout="centered")

st.title("📋 중량물 작업계획서 생성기")

fields = {
    "작업명": "",
    "작업일자": datetime.today().strftime('%Y-%m-%d'),
    "작업 위치": "",
    "중량물 명칭": "",
    "중량": "",
    "크기": "",
    "사용 장비": "",
    "작업 책임자": "",
    "위험요소": "",
    "안전조치": ""
}

data = {}

with st.form("form"):
    for key in fields:
        data[key] = st.text_input(key, fields[key])
    submitted = st.form_submit_button("📁 엑셀 파일 생성")

if submitted:
    # 엑셀 생성
    wb = Workbook()
    ws = wb.active
    ws.title = "중량물 작업계획서"

    header_font = Font(bold=True, size=12)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("A1:B1")
    ws["A1"] = "중량물 작업계획서"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = center_align

    row = 3
    for key, value in data.items():
        ws[f"A{row}"] = key
        ws[f"B{row}"] = value
        ws[f"A{row}"].font = header_font
        ws[f"A{row}"].alignment = center_align
        ws[f"B{row}"].alignment = Alignment(wrap_text=True)
        row += 1

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 50

    # 엑셀 저장 → 메모리 버퍼
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    # 파일 다운로드 버튼
    st.success("✅ 엑셀 파일이 생성되었습니다!")
    st.download_button(
        label="📥 엑셀 다운로드",
        data=buffer,
        file_name=f"중량물작업계획서_{data['작업일자']}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
