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
import numpy as np # pandasが依存するためインポートを明記

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
            df = pd.read_csv(io.BytesIO(raw_data), header=1, encoding=encoding)
            
            if '年' in df.columns:
                 return df
            else:
                 continue

        except Exception:
            continue
            
    raise UnicodeDecodeError(f"ファイル '{uploaded_file.name}' は、一般的な日本語エンコーディングで読み込めませんでした。")


# --- Excel書き込み関数 (openpyxlで統計値を書き込む) ---
def write_all_data_to_excel(excel_file_path, df_combined, df_before_full, df_after_full, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    """
    Openpyxlを使って、全てのデータとレポート情報をExcelファイルに書き込む。
    """
    SHEET_NAMES = {
        '元データ': df_combined,
        '施工前': df_before_full,
        '施工後（調光後）': df_after_full,
    }
    SHEET1_NAME = 'Sheet1'
    SUMMARY_SHEET_NAME = 'まとめ'
    
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
    except FileNotFoundError:
        st.error(f"エラー: Excelテンプレートが見つかりません。")
        return False

    # --- 1. データシートの上書き (Openpyxlを使用) ---
    for sheet_name, df_data in SHEET_NAMES.items():
        if sheet_name not in workbook.sheetnames:
            workbook.create_sheet(sheet_name)
        ws = workbook[sheet_name]
        
        # 既存のデータをクリア (ヘッダー行を維持したい場合は、2行目から削除)
        # NOTE: CSVは2行目からデータが始まるため、1行目(ヘッダー)を残す
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row) 
        
        # DataFrameを直接セルに書き込む（ヘッダー行は除く）
        rows = dataframe_to_rows(df_data, header=False, index=False)
        for row in rows:
             ws.append(row)

    # --- 2. Sheet1: 24時間別平均の書き込み ---
    if SHEET1_NAME not in workbook.sheetnames:
        workbook.create_sheet(SHEET1_NAME) 
        
    ws_sheet1 = workbook[SHEET1_NAME]
    
    metrics_before = df_before.groupby('時')['合計kWh'].agg(['mean', 'count'])
    metrics_after = df_after.groupby('時')['合計kWh'].agg(['mean', 'count'])

    current_row = 36
    for hour in range(1, 25): 
        ws_sheet1.cell(row=current_row, column=1, value=f"{hour:02d}:00")
        ws_sheet1.cell(row=current_row, column=3, value=metrics_before.loc[hour, 'mean'] if hour in metrics_before.index else 0) 
        ws_sheet1.cell(row=current_row, column=4, value=metrics_after.loc[hour, 'mean'] if hour in metrics_after.index else 0)
        current_row += 1
    
    ws_sheet1['C35'] = '施工前 平均kWh/h'
    ws_sheet1['D35'] = '施工後 平均kWh/h'
    ws_sheet1['A35'] = '時間帯'

    # --- 3. まとめシート: 期間 (H6, H7), 営業時間 (H8), タイトル (B1) の書き込み ---
    if SUMMARY_SHEET_NAME not in workbook.sheetnames:
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

    ws_summary['H6'] = before_str
    ws_summary['H7'] = after_str
    ws_summary['H8'] = operating_hours
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
            df_combined.dropna(subset=['年', '月', '日'], inplace=True)
            
            df_combined['日付'] = pd.to_datetime(
                df_combined['年'].astype(str) + '-' + df_combined['月'].astype(str) + '-' + df_combined['日'].astype(str), 
                format='%Y-%m-%d', errors='coerce'
            ).dt.date
            df_combined.dropna(subset=['日付'], inplace=True)
            
            datetime_cols = ['年', '月', '日', '時', '日付']
            consumption_cols = [col for col in df_combined.columns if col not in datetime_cols and not col.startswith('Unnamed:')]
            
            # 💡 エラー修正: 列データに対して処理を実行
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
            
            # Openpyxlのみで全データを書き込み
            write_all_data_to_excel(temp_excel_path, df_combined, df_before_full, df_after_full, df_before, df_after, start_b, end_b, start_a, end_a, operating_hours, store_name)
            
            
            # --- e) ファイル名の変更とダウンロードの準備 ---
            today_date_str = datetime.date.today().strftime('%Y%m%d')
            new_file_name = f"{store_name}：電力報告書{today_date_str}.xlsx"
            
            final_path = os.path.join(temp_dir, new_file_name)
            os.rename(temp_excel_path, final_path)
            
            # ダウンロードボタンの表示
            with open(final_path, "rb") as file:
                st.success("✅ 処理が完了しました！以下のボタンから報告書をダウンロードしてください。")
                st.download_button(
                    label="⬇️ 報告書ファイルをダウンロード",
                    data=file,
                    file_name=new_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
        except Exception as e:
            st.error("🚨 実行中にエラーが発生しました。ファイル形式と入力値を確認してください。")
            st.warning("特に、CSVのヘッダー行が「年,月,日,時,kWh,...」の形式で2行目から始まっているか確認してください。")
            st.exception(e)

if __name__ == "__main__":
    main_streamlit_app()
