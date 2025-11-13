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


# --- CSV読み込み関数 (エンコーディング自動検出) ---
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
            # ヘッダー行をスキップ (header=1, 0-indexed)
            df = pd.read_csv(io.BytesIO(raw_data), header=1, encoding=encoding) 
            
            # 必要なカラム名 '年' が存在するかで成功を判断
            if '年' in df.columns:
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
        ws_sheet1.cell(row=current_row, column=2, value=time_range) 
        
        # C列 (施工前 平均)
        value_before = 0.0
        if metrics_before is not None and start_hour in metrics_before.index:
            mean_val = metrics_before.loc[start_hour, 'mean']
            value_before = float(mean_val) if not np.isnan(mean_val) else 0.0
        ws_sheet1.cell(row=current_row, column=3, value=value_before)
            
        # D列 (施工後 平均)
        value_after = 0.0
        if metrics_after is not None and start_hour in metrics_after.index:
            mean_val = metrics_after.loc[start_hour, 'mean']
            value_after = float(mean_val) if not np.isnan(mean_val) else 0.0
        ws_sheet1.cell(row=current_row, column=4, value=value_after)
            
        current_row += 1
    
    # シートのヘッダーがもし上書きされていなければ設定（テンプレートに依存）
    ws_sheet1['C35'] = '施工前 平均kWh/h'
    ws_sheet1['D35'] = '施工後 平均kWh/h'
    ws_sheet1['A35'] = '時間帯'

    # --- 2. まとめシート: 期間 (H6, H7), 営業時間 (H8), タイトル (B1), 合計値 (B7, B8) の書き込み ---
    if SUMMARY_SHEET_NAME not in workbook.sheetnames:
        workbook.create_sheet(SUMMARY_SHEET_NAME)
        
    ws_summary = workbook[SUMMARY_SHEET_NAME]

    format_date = lambda d: f"{d.year}/{d.month}/{d.day}"

    start_b_str = format_date(start_before)
    end_b_str = format_date(end_before)
    before_str = f"施工前：{start_b_str}～{end_b_str}（{days_before}日間）"
    
    start_a_str = format_date(start_after)
    end_a_str = format_date(end_after)
    after_str = f"施工後(調光後)：{start_a_str}～{end_a_str}（{days_after}日間）"

    ws_summary['H6'] = before_str
    ws_summary['H7'] = after_str
    ws_summary['H8'] = operating_hours
    ws_summary['B1'] = f"{store_name}の使用電力比較報告書"
    
    # まとめシートの合計値も書き込み (日別平均合計kWh)
    ws_summary['B7'] = float(avg_daily_total_before)
    ws_summary['B8'] = float(avg_daily_total_after)
    
    # ファイルを保存
    workbook.save(excel_file_path)
    
    return True


# --- Streamlitメインアプリケーション ---
def main_streamlit_app():
    st.set_page_config(layout="wide", page_title="電力データ報告書作成アプリ")
    st.title("💡 電力データ自動処理アプリ")
    st.markdown("### Step 1: ファイルのアップロード")
    
    # --- 1. CSVファイルのアップロード ---
    uploaded_csvs = st.file_uploader(
        "📈 CSVデータ (複数可) をアップロードしてください",
        type=['csv'],
        accept_multiple_files=True
    )
    
    if uploaded_csvs:
        st.success(f"CSVファイル {len(uploaded_csvs)}個 が準備できました。")
        st.markdown("---")
        st.markdown("### Step 2: 期間と情報の入力")
    else:
        st.warning("処理を開始するには、CSVデータをアップロードしてください。")
        return

    # --- 2. ユーザー入力ウィジェット ---
    today =
