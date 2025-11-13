import streamlit as st
import pandas as pd
import os
import sys
import chardet
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime
import shutil
import io
import numpy as np

# ======================================================
# 💡 設定: ファイル名
# ======================================================
EXCEL_TEMPLATE_FILENAME = '富士川店：電力報告250130.xlsx'


# ======================================================
# 📂 CSV読み込み関数 (自動エンコーディング検出)
# ======================================================
@st.cache_data
def detect_and_read_csv(uploaded_file):
    """アップロードされたファイルの内容を読み込み、エンコーディングを自動検出してDataFrameを返す"""
    
    uploaded_file.seek(0)
    raw_data = uploaded_file.read()
    
    detected_encoding = chardet.detect(raw_data)['encoding']
    encodings_to_try = ['cp932', 'shift_jis', 'utf-8']
    
    if detected_encoding and detected_encoding.lower() not in encodings_to_try:
        encodings_to_try.append(detected_encoding.lower())

    for encoding in encodings_to_try:
        try:
            # header=1 (2行目) をヘッダーとして読み込む
            df = pd.read_csv(io.Bytes_
