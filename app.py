import streamlit as st
import pandas as pd
import os
import glob
import sys
import chardet
import openpyxl
from openpyxl.utils import cell
import datetime
import shutil
import io

# ======================================================
# 💡 設定: ファイル名
# ======================================================
# GitHubリポジトリに置くテンプレートExcelファイル名
EXCEL_TEMPLATE_FILENAME = '富士川店：電力報告250130.xlsx'


# --- CSV読み込み関数 (自動エンコーディング検出) ---
@st.cache_data
def detect_and_read_csv(uploaded_file):
    """アップロードされたファイルの内容を読み込み、エンコーディングを自動検出してDataFrameを返す"""
    
    # ファイルポインタをリセット
    uploaded_file.seek(0)
    raw_data = uploaded_file.read()
    
    detected_encoding = chardet.detect(raw_data)['encoding']
    encodings_to_try = ['cp932', 'shift_jis', 'utf-8']
    
    if detected_encoding and detected_encoding.lower() not in encodings_to_try:
        encodings_to_try.append(detected_encoding.lower())

    for encoding in encodings_to_try:
        try:
            # BytesIOを使ってメモリから読み込み
            df = pd.read_csv(io.BytesIO(raw_data), header=1, encoding=encoding)
            
            # データフレームに「年」カラムがあることをチェック
            if '年' in df.columns:
                 return df
            else:
                 continue

        except Exception:
            continue
            
    raise UnicodeDecodeError(f"ファイル '{uploaded_file.name}' は、一般的な日本語エンコーディングで読み込めませんでした。")


# --- Excel書き込み関数 (openpyxlで統計値を書き込む) ---
def write_excel_reports(excel_file_path, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    """
    Excelファイルにデータ、平均値、期間情報を書き込む。
    """
    SHEET1_NAME = 'Sheet1'
    SUMMARY_SHEET_NAME = 'まとめ'
    
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
    except FileNotFoundError:
        st.error(f"エラー: Excelテンプレートが見つかりません。")
        return

    # --- 1. Sheet1: 24時間別平均の書き込み (C36～D59) ---
    if SHEET1_NAME not in workbook.sheetnames:
        # Sheet1がない場合は作成（通常はテンプレートにあるはず）
        workbook.create_sheet(SHEET1_NAME) 
        
    ws_sheet1 = workbook[SHEET1_NAME]
    
    # 時間帯別平均の計算
    metrics_before = df_before.groupby('時')['合計kWh'].agg(['mean', 'count'])
    metrics_after = df_after.groupby('時')['合計kWh'].agg(['mean', 'count'])

    current_row = 36
    for hour in range(1, 25): 
        # 時刻ラベル (A列)
        ws_sheet1.cell(row=current_row, column=1, value=f"{hour:02d}:00")
        
        # 施工前平均 (C列)
        ws_sheet1.cell(row=current_row, column=3, 
                       value=metrics_before.loc[hour, 'mean'] if hour in metrics_before.index else 0) 
        
        # 施工後平均 (D列)
        ws_sheet1.cell(row=current_row, column=4, 
                       value=metrics_after.loc[hour, 'mean'] if hour in metrics_after.index else 0)
        current_row += 1
    
    # ヘッダー (35行目)
    ws_sheet1['C35'] = '施工前 平均kWh/h'
    ws_sheet1['D35'] = '施工後 平均kWh/h'
    ws_sheet1['A35'] = '時間帯'

    # --- 2. まとめシート: 期間 (H6, H7), 営業時間 (H8), タイトル (B1) の書き込み ---
    if SUMMARY_SHEET_NAME not in workbook.sheetnames:
        # まとめシートがない場合は作成
        workbook.create_sheet(SUMMARY_SHEET_NAME)
        
    ws_summary = workbook[SUMMARY_SHEET_NAME]

    days_before = (end_before - start_before).days + 1
    days_after = (end_after - start_after).days + 1
    format_date = lambda d: f"{d.year}/{d.month}/{d.day}"

    start_b_str = format_date(start_before)
    end_b_str = format_date(end_before)
    before_str = f"施工前：{start_b_str}～{end_b_str}（{days_before}日間）"
    
    start_a_str = format_date(start_after)
    end_a_str = format_date(end_after)
    after_str = f"施工後(調光後)：{start_a_str}～{end_a_str}（{days_after}日間）"

    # H列への書き込み
    ws_summary['H6'] = before_str
    ws_summary['H7'] = after_str
    ws_summary['H8'] = operating_hours
    
    # B1へのタイトル書き込み
    ws_summary['B1'] = f"{store_name}の使用電力比較報告書"
    
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
        end_before = st.date_input("終了日", today - datetime.timedelta(days=25), key="end_b")
        
    with col_date2:
        st.subheader("📅 施工後 測定期間")
        start_after = st.date_input("開始日", today - datetime.timedelta(days=10), key="start_a")
        end_after = st.date_input("終了日", today - datetime.timedelta(days=5), key="end_a")

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
            # テンプレートファイルをテンポラリフォルダにコピー
            shutil.copy(EXCEL_TEMPLATE_FILENAME, temp_excel_path)
                
            # --- b) データ統合と前処理 ---
            all_data = []
            for csv_file in uploaded_csvs:
                # CSVの読み込みとエンコーディング検出
                df = detect_and_read_csv(csv_file)
                all_data.append(df)
            df_combined = pd.concat(all_data, ignore_index=True)
            
            # データ前処理（日付の結合と合計kWhの計算）
            df_combined['年'] = pd.to_numeric(df_combined['年'], errors='coerce').astype('Int64')
            df_combined['月'] = pd.to_numeric(df_combined['月'], errors='coerce').astype('Int64')
            df_combined['日'] = pd.to_numeric(df_combined['日'], errors='coerce').astype('Int64')
            df_combined.dropna(subset=['年', '月', '日'], inplace=True)
            
            # 日付の結合
            df_combined['日付'] = pd.to_datetime(
                df_combined['年'].astype(str) + '-' + df_combined['月'].astype(str) + '-' + df_combined['日'].astype(str), 
                format='%Y-%m-%d', errors='coerce'
            ).dt.date
            df_combined.dropna(subset=['日付'], inplace=True)
            
            # 合計kWhの計算
            datetime_cols = ['年', '月', '日', '時', '日付']
            consumption_cols = [col for col in df_combined.columns if col not in datetime_cols and not col.startswith('Unnamed:')]
            for col in consumption_cols:
                df_combined[col] = pd.to_numeric(df_combined[col], errors='coerce').fillna(0)
            df_combined['合計kWh'] = df_combined[consumption_cols].sum(axis=1)


            # --- c) データ分割 ---
            start_b = start_before
            end_b = end_before
            start_a = start_after
            end_a = end_after

            df_before_full = df_combined[(df_combined['日付'] >= start_b) & (df_combined['日付'] <= end_b)].copy()
            df_after_full = df_combined[(df_combined['日付'] >= start_a) & (df_combined['日付'] <= end_a)].copy()
            df_before = df_before_full.copy()
            df_after = df_after_full.copy()
            
            # --- d) Excel書き込み ---
            
            # 1. Pandasでデータシートを上書き
            # 💡 最新のPandasで既存シートを保持しながら書き込む方法を採用
            
            # 既存のワークブックを読み込む
            existing_workbook = openpyxl.load_workbook(temp_excel_path)
            
            with pd.ExcelWriter(temp_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # 既存のワークブックオブジェクトをwriterにセット
                writer.book = existing_workbook
                
                # 既存のシート群をwriterに登録し、Pandasが管理していることを示す
                # これにより、Pandasが書き込まないシート（Sheet1, まとめ）は保持される
                writer.sheets = dict((ws.title, ws) for ws in existing_workbook.worksheets)

                # データシートの上書き（データシートのみif_sheet_exists='replace'で上書き）
                df_combined.to_excel(writer, sheet_name='元データ', index=False) 
                df_before_full.to_excel(writer, sheet_name='施工前', index=False)   
                df_after_full.to_excel(writer, sheet_name='施工後（調光後）', index=False)

            # 2. OpenPyXLで統計値シート（Sheet1, まとめ）を更新
