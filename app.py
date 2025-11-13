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
            header_list = df_full.iloc[header_row_index].tolist()
            
            # '年', '月', '日', '時' の後のカラムを 'kWh_1', 'kWh_2', ... と命名し直す
            cleaned_columns = []
            kWh_counter = 1
            for i, col in enumerate(header_list):
                # 最初の4列（A, B, C, D）を固定
                if i < 4:
                    cleaned_columns.append(col)
                # 5列目以降 (E列以降) を電力消費データとして扱う
                elif i >= 4:
                    cleaned_columns.append(f'kWh_{kWh_counter}')
                    kWh_counter += 1
                else:
                    # これは発生しないはずだが、念のため
                    cleaned_columns.append(f'Unnamed_{i}')

            df.columns = cleaned_columns

            # データが存在し、必要なカラム名 '年' が存在するかで成功を判断
            if '年' in df.columns and not df.empty:
                return df
            else:
                continue

        except Exception:
            continue
            
    raise Exception(f"ファイル '{uploaded_file.name}' は、一般的な日本語エンコーディングで読み込めませんでした。")


# --- Excelレポート書き込み関数 ---
def write_excel_reports(excel_file_path, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    
    SHEET1_NAME = 'Sheet1'
    SUMMARY_SHEET_NAME = 'まとめ'
    
    try:
        # テンプレートExcelファイルを読み込む
        workbook = openpyxl.load_workbook(excel_file_path)
    except FileNotFoundError:
        st.error(f"エラー: Excelテンプレートが見つかりません。")
        return False

    # --- 共通計算 ---
    days_before = (end_before - start_before).days + 1
    days_after = (end_after - start_after).days + 1
    
    # 測定期間中の日別平均合計kWhを計算 (合計kWhを総日数で割る)
    total_kWh_before = df_before['合計kWh'].sum()
    total_kWh_after = df_after['合計kWh'].sum()
    
    # NaN/ZeroDivision チェック
    avg_daily_total_before = total_kWh_before / days_before if days_before > 0 and not np.isnan(total_kWh_before) else 0
    avg_daily_total_after = total_kWh_after / days_after if days_after > 0 and not np.isnan(total_kWh_after) else 0
    
    # --- 1. Sheet1: 24時間別平均の書き込み (C36～D59) と合計値 (C33, D33) ---
    if SHEET1_NAME not in workbook.sheetnames:
        workbook.create_sheet(SHEET1_NAME) 
        
    ws_sheet1 = workbook[SHEET1_NAME]
    
    # C33, D33に日別平均合計値を書き込む
    ws_sheet1['C33'] = float(avg_daily_total_before)
    ws_sheet1['D33'] = float(avg_daily_total_after)
    
    # 24時間別平均の計算（「時」カラムは0-23に標準化済み）
    metrics_before = df_before.groupby('時')['合計kWh'].agg(['mean', 'count']) if not df_before.empty else None
    metrics_after = df_after.groupby('時')['合計kWh'].agg(['mean', 'count']) if not df_after.empty else None

    current_row = 36
    # 0時から23時までの開始時間でループ (これがグループキーとなる)
    for start_hour in range(0, 24):
        
        # 時間帯の表示
        end_hour = (start_hour + 1) % 24
        time_range = f"{start_hour:02d}:00～{end_hour:02d}:00"

        # A列: 内部ID（00:00, 01:00...）
        ws_sheet1.cell(row=current_row, column=1, value=f"{start_hour:02d}:00") 
        # B列: 時間帯表記
        ws_sheet1.cell(row=current
