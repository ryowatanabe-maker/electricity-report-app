import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import shutil
import chardet
import datetime
import openpyxl
import matplotlib.pyplot as plt

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
    # Ensure raw is bytes for chardet
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
            # 1. ヘッダー指定なしでファイル全体を読み込む
            df_full = pd.read_csv(io.BytesIO(raw), header=None, encoding=enc, keep_default_na=False)
            header_row_index = -1
            
            # 2. ヘッダー行（'年', '月', '日', '時' を含む行）を特定する
            for i in range(df_full.shape[0]):
                # 必須カラムが含まれるかチェック
                row = df_full.iloc[i].astype(str).tolist()
                if all(x in row for x in ['年', '月', '日', '時']):
                    header_row_index = i
                    break
            
            if header_row_index == -1:
                continue

            # 3. 実際のデータ行を抽出し、カラム名を標準化する
            header = df_full.iloc[header_row_index].tolist()
            data = df_full.iloc[header_row_index + 1:].copy().reset_index(drop=True)

            # カラム名整形：A-D はそのまま、E以降は kWh_1...
            cleaned_cols = []
            k = 1
            for i, col in enumerate(header):
                if i < 4:
                    cleaned_cols.append(str(col))
                else:
                    cleaned_cols.append(f'kWh_{k}')
                    k += 1

            # 読み込んだ行数とヘッダー長がずれる場合の補正
            if data.shape[1] != len(cleaned_cols):
                while len(cleaned_cols) < data.shape[1]:
                    cleaned_cols.append(f'Unnamed_{len(cleaned_cols)}')
                if len(cleaned_cols) > data.shape[1]:
                    cleaned_cols = cleaned_cols[:data.shape[1]]

            data.columns = cleaned_cols

            # 必須カラムチェック
            if not all(col in data.columns for col in ['年', '月', '日', '時']):
                continue

            return data

        except Exception:
            continue

    raise Exception(f"CSVファイル '{getattr(uploaded_file, 'name', 'unknown')}' を適切に読み込めませんでした（エンコーディング/形式を確認してください）。")

# ---------------------------
# Excel書き込み関数
# ---------------------------
def write_excel_reports(excel_path, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    """
    - 0-23 時毎の平均を算出し、Sheet1 に C36-C59 (before), D36-D59 (after) として書き込む
    - まとめシートに期間・営業時間・店舗名を書き込む
    """
    SHEET1 = "Sheet1"
    SUMMARY = "まとめ"

    try:
        wb = openpyxl.load_workbook(excel_path)
    except FileNotFoundError:
        st.error("Excelテンプレートが見つかりません。")
        return False

    # Helper to calculate hourly mean for the period
    def hourly_mean_series(df):
        if df is None or df.empty:
            # データがない場合は 0.0 のシリーズを返す
            return pd.Series([0.0]*24, index=range(24), dtype=float)
        
        # '時'でグループ化し、'合計kWh'の平均を計算
        ser = df.groupby('時')['合計kWh'].mean()
        ser.index = ser.index.astype(int)
        # 0-23時のインデックスに再調整し、欠損時間を 0.0 で埋める
        ser = ser.reindex(range(24), fill_value=0.0)
        return ser

    ser_before = hourly_mean_series(df_before)
    ser_after
