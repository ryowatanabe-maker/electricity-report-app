import streamlit as st
import pandas as pd
import os
import glob
import sys
import chardet
import openpyxl
from openpyxl.utils import cell
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime
import shutil
import io
import numpy as np

# ======================================================
# 💡 設定: ファイル名
# ======================================================
EXCEL_TEMPLATE_FILENAME = '電力報告テンプレート.xlsx'


# --- CSV読み込み関数 (エンコーディング自動検出 & ヘッダー処理修正) ---
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
            # ヘッダー指定なし (header=None) でファイル全体を読み込む
            df_full = pd.read_csv(io.BytesIO(raw_data), header=None, encoding=encoding, keep_default_na=False) 
            
            # ヘッダーとして使用する行（年,月,日,時,...の行）を特定
            header_row_index = -1
            if not df_full.empty:
                 # '年'を含む行を探し、それをヘッダー行とする
                for i in range(df_full.shape[0]):
                    # 最初の4カラムに '年', '月', '日', '時' が含まれているかチェック
                    # CSVファイルがUTF-8 BOMやその他の文字を含む可能性があるため、astype(str)で安全に比較
                    row_values = df_full.iloc[i].astype(str).tolist()
                    if '年' in row_values and '月' in row_values and '日' in row_values and '時' in row_values:
                        header_row_index = i
                        break
            
            if header_row_index == -1:
                 # ヘッダーが見つからなかった場合はスキップして次のエンコーディングへ
                 continue

            # 実際のデータ行を抽出 (ヘッダー行の次から)
            df = df_full.iloc[header_row_index + 1:].copy()
            
            # 💡 カラム名の再設定ロジック
            # 読み込み時のカラム名をリストとして取得
            header_list = df_full.iloc[header_row_index].tolist()
            
            # '年', '月', '日', '時' の後のカラムを 'kWh_1', 'kWh_2', ... と命名し直す
            cleaned_columns = []
            kWh_counter = 1
            for i, col in enumerate(header_list):
                if i < 4:
                    # 最初の4列（A, B, C, D）を固定
                    cleaned_columns.append(col)
                elif i >= 4:
                    # 5列目以降 (E列以降) を電力消費データとして扱う
                    cleaned_columns.append(f'kWh_{kWh_counter}')
                    kWh_counter += 1
                else:
                    # 想定外のカラム名（予備）
                    cleaned_columns.append(f'Unnamed_{i}')

            df.columns = cleaned_columns

            # データが存在し、必要なカラム名 '年' が存在するかで成功を判断
