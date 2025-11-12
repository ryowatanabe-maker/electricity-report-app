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
            # 💡 修正: header=0 (1行目) をヘッダーとして読み込む
            df = pd.read_csv(io.BytesIO(raw_data), header=0, encoding=encoding)
            
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
    # 💡 修正: 0時から23時までループ (合計24時間分)
    for hour in range(0, 24): 
        
        # CSVの '時' カラムの値は 1-24 または 0-23 のどちらかの可能性あり。
        # 0:00 のデータは CSV上は '時'=0 または '時'=24 であるため、両方を考慮
        
        # CSVの '時'カラムが 1-24 の場合: hour+1
        # CSVの '時'カラムが 0-23 の場合: hour
        
        # 両方に対応するため、hour (0-23) をキーとして使用し、0時と24時(翌日0時)を区別せず集計します。
        
        # テンプレートに合わせた時間帯ラベルの計算 (例: 00:00～01:00)
        display_hour = (hour + 1) % 24
        if display_hour == 0:
            display_hour = 24 # 24時として表示
            
        start_h = f"{hour:02d}:00"
        end_h = f"{display_hour:02d}:00"
        time_range = f"{start_h}～{end_h}"
        
        ws_sheet1.cell(row=current_row, column=1, value=f"{hour:02d}") # A列に00, 01, ...
        ws_sheet1.cell(row=current_row, column=2, value=time_range) 
        
        # C列 (施工前 平均)
        # 💡 CSVの '時' カラムが 1-24 の場合と 0-23 の場合の両方に対応
        mean_b = 0
        if metrics_before is not None:
             if hour in metrics_before.index: # 0-23時形式の場合
                 mean_b = metrics_before.loc[hour, 'mean']
             elif hour + 1 in metrics_before.index: # 1-24時形式の場合 (例: 0時データは24時として記録)
                 mean_b = metrics_before.loc[hour + 1, 'mean']
        ws_sheet1.cell(row=current_row, column=3, value=mean_b)

        # D列 (施工後 平均)
        mean_a = 0
        if metrics_after is not None:
             if hour in metrics_after.index:
                 mean_a = metrics_after.loc[hour, 'mean']
             elif hour + 1 in metrics_after.index:
                 mean_a = metrics_after.loc[hour + 1, 'mean']
        ws_sheet1.cell(row=current_row, column=4, value=mean_a)
             
        current_row += 1
    
    ws_sheet1['C35'] = '施工前 平均kWh/h'
    ws_sheet1['D35'] = '施工後 平均kWh/h'
    ws_sheet1['A35'] = '時' # 時刻のインデックスを示す
    ws_sheet1['B35'] = '時間帯'


    # --- 2. まとめシート: 期間 (H6, H7), 営業時間 (H8), タイトル (B1) の書き込み ---
    if SUMMARY_SHEET_NAME not in workbook.sheetnames:
        workbook.create_sheet(SUMMARY_SHEET_NAME)
        
    ws_summary = workbook[SUMMARY_SHEET_NAME]

    format_date = lambda d: f"{d.year}/{d.month}/{d.day}"

    days_before = (end_before - start_before).days + 1
    days_after = (end_after - start_after).days + 1

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
    
    # まとめシートの合計値も書き込み (B7, B8を推定)
    ws_summary['B7'] = avg_daily_total_before
    ws_summary['B8'] = avg_daily_total_after
    
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
    today = datetime.date.today()
    
    col_date1, col_date2 = st.columns(2)
    
    with col_date1:
        st.subheader("🗓️ 施工前 測定期間")
        start_before = st.date_input("開始日", today - datetime.timedelta(days=30), key="start_b")
        end_before = st.date_input("終了日", today - datetime.timedelta(days=23), key="end_b")
        
    with col_date2:
        st.subheader("📅 施工後 測定期間")
        start_after = st.date_input("開始日", today - datetime.timedelta(days=14), key="start_a")
        end_after = st.date_input("終了日", today - datetime.timedelta(days=7), key="end_a")

    col_info1, col_info2 = st.columns(2)
    with col_info1:
        operating_hours = st.text_input("営業時間", value="08:00-22:00", help="まとめシートH8に反映")
    with col_info2:
        store_name = st.text_input("店舗名", value="大倉山店", help="報告書名とまとめシートB1に反映")
        
    st.markdown("---")
    
    # --- 3. 実行ボタン ---
    if st.button("🚀 データ処理を実行し、報告書をダウンロード"):
        try:
            # テンポラリフォルダのセットアップ
            temp_dir = "temp_data"
            os.makedirs(temp_dir, exist_ok=True)
            
            # --- a) テンプレートExcelファイルをGitHubからコピー ---
            if not os.path.exists(EXCEL_TEMPLATE_FILENAME):
                 st.error(f"🚨 致命的なエラー: GitHubリポジトリにテンプレートファイル '{EXCEL_TEMPLATE_FILENAME}' が見つかりません。ファイル名を確認し、app.pyと同じ場所に配置してください。")
                 return

            temp_excel_path = os.path.join(temp_dir, EXCEL_TEMPLATE_FILENAME)
            shutil.copy(EXCEL_TEMPLATE_FILENAME, temp_excel_path)
                
            # --- b) データ統合と前処理 ---
            all_data = []
            for csv_file in uploaded_csvs:
                df = detect_and_read_csv(csv_file)
                all_data.append(df)
            df_combined = pd.concat(all_data, ignore_index=True)
            
            # データ前処理（日付の結合と合計kWhの計算）
            df_combined['年'] = pd.to_numeric(df_combined['年'], errors='coerce').astype('Int64')
            df_combined['月'] = pd.to_numeric(df_combined['月'], errors='coerce').astype('Int64')
            df_combined['日'] = pd.to_numeric(df_combined['日'], errors='coerce').astype('Int64')
            
            # --- データの重複削除 (同一日時レコードの削除) ---
