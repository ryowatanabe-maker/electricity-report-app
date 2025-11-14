# app.py
import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import shutil
import chardet
import datetime
import openpyxl

# ---------------------------
# 設定
# ---------------------------
EXCEL_TEMPLATE_FILENAME = "電力報告テンプレート.xlsx"
TEMP_DIR = "temp_data"

# ---------------------------
# ヘッダー自動検出 + CSV読み込み
# ---------------------------
@st.cache_data
def detect_and_read_csv(uploaded_file) -> pd.DataFrame:
    """
    アップロードCSVのエンコーディングを検出し、'年','月','日','時' を含むヘッダー行を探してDataFrameを返す。
    E列以降は kWh_1, kWh_2 ... としてリネームする。
    """
    uploaded_file.seek(0)
    raw = uploaded_file.read()
    if isinstance(raw, str):
        raw = raw.encode('utf-8')
    detect = chardet.detect(raw)
    encodings_to_try = ['cp932', 'shift_jis', 'utf-8']
    if detect and detect.get('encoding'):
        enc = detect['encoding'].lower()
        if enc not in encodings_to_try:
            encodings_to_try.append(enc)

    for enc in encodings_to_try:
        try:
            # まず全体をヘッダーなしで読み込み（バイナリから）
            df_full = pd.read_csv(io.BytesIO(raw), header=None, encoding=enc, keep_default_na=False)
            # ヘッダー行を探す（'年','月','日','時' を含む行）
            header_row_index = -1
            for i in range(df_full.shape[0]):
                row = df_full.iloc[i].astype(str).tolist()
                if all(x in row for x in ['年', '月', '日', '時']):
                    header_row_index = i
                    break
            if header_row_index == -1:
                continue

            header = df_full.iloc[header_row_index].tolist()
            data = df_full.iloc[header_row_index + 1:].copy().reset_index(drop=True)

            # カラム名整形：A-D はそのまま、E以降は kWh_1...
            cleaned_cols = []
            k = 1
            for i, col in enumerate(header):
                if i < 4:
                    cleaned_cols.append(str(col))
                else:
