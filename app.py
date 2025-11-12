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
EXCEL_TEMPLATE_FILENAME = '富士川店：電力報告250130.xlsx'


# --- CSV読み込み関数 (自動エンコーディング検出) ---
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
            # 💡 header=1 (2行目) をヘッダーとして読み込む設定に戻す (CSV構造に合わせるため)
            df = pd.read_csv(io.BytesIO(raw_data), header=1, encoding=encoding)
            
            if '年' in df.columns:
                 return df
            else:
                 continue

        except Exception:
            continue
            
    raise Exception(f"ファイル '{uploaded_file.name}' は、一般的な日本語エンコーディングで読み込めませんでした。")


# --- Excelレポート書き込み関数 (Openpyxlで統計値を書き込む) ---
def write_excel_reports(excel_file_path, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    """
    Openpyxlを使って、Sheet1とまとめシートにレポート情報を書き込む。
    """
    SHEET1_NAME = 'Sheet1'
    SUMMARY_SHEET_NAME = 'まとめ'
    
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
    except FileNotFoundError:
        st.error(f"エラー: Excelテンプレートが見つかりません。")
        return False

    # --- 共通計算 ---
    days_before = (end_before - start_before).days + 1
    days_after = (end_after - start_after).days + 1
    
    # 測定期間中の日別平均合計kWhを計算 (合計kWhを総日数で割る)
    avg_daily_total_before = df_before['合計kWh'].sum() / days_before if not df_before.empty else 0
    avg_daily_total_after = df_after['合計kWh'].sum() / days_after if not df_after.empty else 0
    
    
    # --- 1. Sheet1: 24時間別平均の書き込み (C36～D59) と合計値 (C33, D33) ---
    if SHEET1_NAME not in workbook.sheetnames:
        workbook.create_sheet(SHEET1_NAME) 
        
    ws_sheet1 = workbook[SHEET1_NAME]
    
    # C33, D33に日別平均合計値を書き込む
    ws_sheet1['C33'] = avg_daily_total_before
    ws_sheet1['D33'] = avg_daily_total_after
    
    # 24時間別平均の計算と書き込み
    metrics_before = df_before.groupby('時')['合計kWh'].agg(['mean', 'count']) if not df_before.empty else None
    metrics_after = df_after.groupby('時')['合計kWh'].agg(['mean', 'count']) if not df_after.empty else None

    current_row = 36
    for hour in range(1, 25): 
        # A列: 時間ラベル (e.g., "01:00")
        ws_sheet1.cell(row=current_row, column=1, value=f"{hour:02d}:00") 
        
        # B列: 時間帯ラベル (e.g., "00:00～01:00")
        start_h_val = (hour - 1) % 24
        end_h_val = hour % 24
        start_h = f"{start_h_val:02d}:00"
        end_h = f"{end_h_val:02d}:00"
        time_range = f"{start_h}～{end_h}"

        ws_sheet1.cell(row=current_row, column=2, value=time_range) 
        
        # C列 (施工前 平均)
        if metrics_before is not None and hour in metrics_before.index:
             ws_sheet1.cell(row=current_row, column=3, value=metrics_before.loc[hour, 'mean'])
        else:
             ws_sheet1.cell(row=current_row, column=3, value=0)
             
        # D列 (施工後 平均)
        if metrics_after is not None and hour in metrics_after.index:
             ws_sheet1.cell(row=current_row, column=4, value=metrics_after.loc[hour, 'mean'])
        else:
             ws_sheet1.cell(row=current_row, column=4, value=0)
             
        current_row += 1
    
    ws_sheet1['C35'] = '施工前 平均kWh/h'
    ws_sheet1['D35'] = '施工後 平均kWh/h'
    ws_sheet1['A35'] = '時間帯'

    # --- 2. まとめシート: 期間 (H6, H7), 営業時間 (H8), タイトル (B1) の書き込み ---
    if SUMMARY_SHEET_NAME not in workbook.sheetnames:
        workbook.create_sheet(SUMMARY_SHEET_NAME)
