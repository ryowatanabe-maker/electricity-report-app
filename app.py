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


# --- CSV読み込み関数 (変更なし) ---
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
        workbook = openpyxl.load_workbook(excel_file_path)
    except FileNotFoundError:
        st.error(f"エラー: Excelテンプレートが見つかりません。")
        return False

    # --- 共通計算 ---
    days_before = (end_before - start_before).days + 1
    days_after = (end_after - start_after).days + 1
    
    # 【変更なし】測定期間中の日別平均合計kWhを計算 (合計kWhを総日数で割る)
    # これが「まとめ」シートのB7, B8および「Sheet1」のC33, D33に書き込まれる値です。
    # これは (全期間の合計kWh) / (期間の日数) であり、期間中の日々の平均総消費電力を示します。
    total_kWh_before = df_before['合計kWh'].sum()
    total_kWh_after = df_after['合計kWh'].sum()
    
    # NaNチェック
    avg_daily_total_before = total_kWh_before / days_before if days_before > 0 and not np.isnan(total_kWh_before) else 0
    avg_daily_total_after = total_kWh_after / days_after if days_after > 0 and not np.isnan(total_kWh_after) else 0
    
    # --- 1. Sheet1: 24時間別平均の書き込み (C36～D59) と合計値 (C33, D33) ---
    if SHEET1_NAME not in workbook.sheetnames:
        workbook.create_sheet(SHEET1_NAME) 
        
    ws_sheet1 = workbook[SHEET1_NAME]
    
    # C33, D33に日別平均合計値を書き込む
    ws_sheet1['C33'] = float(avg_daily_total_before)
    ws_sheet1['D33'] = float(avg_daily_total_after)
    
    # 24時間別平均の計算
    # 【ご要望反映】時間帯ごとにグルーピングし、「合計kWh」の平均値を算出
    # これは、期間中の同じ時間帯（例：10時台）の平均消費電力を示します。
    # pandasはNaNを含む行を自動で無視して平均を計算します。
    metrics_before = df_before.groupby('時')['合計kWh'].mean()
    metrics_after = df_after.groupby('時')['合計kWh'].mean()

    current_row = 36
    for hour in range(1, 25): # hourは1から24まで
        # CSVデータによっては「時」が1-24（例：24=0時台）または0-23（例：0=0時台）の場合があるため、1-24で処理
        
        # 見出しの設定
        start_h_val = (hour - 1) % 24
        end_h_val = hour % 24
        start_h = f"{start_h_val:02d}:00"
        end_h = f"{end_h_val:02d}:00"
        time_range = f"{start_h}～{end_h}"

        # A列: 内部IDとして使用（Excelの計算式には影響しない）
        ws_sheet1.cell(row=current_row, column=1, value=f"{hour:02d}:00") 
        # B列: 時間帯表記
        ws_sheet1.cell(row=current_row, column=2, value=time_range) 
        
        # C列 (施工前 平均)
        # hourがmetricsのインデックスにあればその平均値を、なければ0をセット
        value_before = metrics_before.get(hour, 0)
        # NaNチェックをして0.0を書き込む
        ws_sheet1.cell(row=current_row, column=3, value=float(value_before) if not np.isnan(value_before) else 0.0)
            
        # D列 (施工後 平均)
        value_after = metrics_after.get(hour, 0)
        # NaNチェックをして0.0を書き込む
        ws_sheet1.cell(row=current_row, column=4, value=float(value_after) if not np.isnan(value_after) else 0.0)
            
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
    today = datetime.date.today()
    
    col_date1, col_date2 = st.columns(2)
    
    with col_date1:
        st.subheader("🗓️ 施工前 測定期間")
        # デフォルト値を少し現実に合わせて変更
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
        # 期間のバリデーション
        if start_before >= end_before or start_after >= end_after:
            st.error("🚨 期間の設定が不正です。開始日は終了日よりも前に設定してください。")
            return

        try:
            # テンポラリフォルダのセットアップ
            temp_dir = "temp_data"
            os.makedirs(temp_dir, exist_ok=True)
            
            # --- a) テンプレートExcelファイルをGitHubからコピー ---
            # NOTE: Streamlit Cloud環境では、このファイルはリポジトリのルートに存在する必要があります。
            if not os.path.exists(EXCEL_TEMPLATE_FILENAME):
                # テンプレートファイルを読み込む代わりに、エラーを出力
                st.error(f"🚨 致命的なエラー: Excelテンプレートファイル '{EXCEL_TEMPLATE_FILENAME}' が実行環境から見つかりません。")
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
            df_combined['時'] = pd.to_numeric(df_combined['時'], errors='coerce').astype('Int64')
            
            # --- データの重複削除 (同一日時レコードの削除) ---
            # これにより、同じ「年/月/日/時」を持つレコードが複数ある場合、最初のもののみが残り、重複合算を防ぎます。
            df_combined.drop_duplicates(subset=['年', '月', '日', '時'], keep='first', inplace=True)
            
            df_combined.dropna(subset=['年', '月', '日', '時'], inplace=True) # 日時カラムにNaNがある行は削除
            
            df_combined['日付'] = pd.to_datetime(
                df_combined['年'].astype(str) + '-' + df_combined['月'].astype('str') + '-' + df_combined['日'].astype('str'), 
                format='%Y-%m-%d', errors='coerce'
            ).dt.date
            df_combined.dropna(subset=['日付'], inplace=True)
            
            datetime_cols = ['年', '月', '日', '時', '日付']
            # E列以降のカラムを消費電力カラムとして特定
            consumption_cols = [col for col in df_combined.columns if col not in datetime_cols and not col.startswith('Unnamed:')]
            
            if not consumption_cols:
                st.error("エラー: E列以降に消費電力データ（kWhや回路データ）のカラムが見つかりませんでした。CSVの形式を確認してください。")
                return

            # 消費電力カラムの数値変換と合算ロジック
            # 【ご要望反映】E列以降の数値を全て合算して「合計kWh」を作成
            for col in consumption_cols:
                df_combined[col] = pd.to_numeric(df_combined[col], errors='coerce').fillna(0)
            
            df_combined['合計kWh'] = df_combined[consumption_cols].sum(axis=1)


            # --- c) データ分割 ---
            start_b = start_before
            end_b = end_before
            start_a = start_after
            end_a = end_after

            # 測定期間内のデータを抽出
            df_before = df_combined[(df_combined['日付'] >= start_b) & (df_combined['日付'] <= end_b)].copy()
            df_after = df_combined[(df_combined['日付'] >= start_a) & (df_combined['日付'] <= end_a)].copy()
            
            # データが空でないか確認
            if df_before.empty:
                st.warning(f"🚨 施工前期間（{start_b}～{end_b}）に対応するデータがアップロードされたCSVファイルに見つかりませんでした。")
            if df_after.empty:
                st.warning(f"🚨 施工後期間（{start_a}～{end_a}）に対応するデータがアップロードされたCSVファイルに見つかりませんでした。")
                
            # --- d) Excel書き込み ---
            
            # OpenPyXLでSheet1とまとめシートを更新（時間帯別平均値と期間情報）
            success = write_excel_reports(temp_excel_path, df_before, df_after, start_b, end_b, start_a, end_a, operating_hours, store_name)
            
            if not success:
                # write_excel_reports内でエラーメッセージが表示されているため、ここでreturn
                return 

            
            # --- e) ファイル名の変更とダウンロードの準備 ---
            today_date_str = datetime.date.today().strftime('%Y%m%d')
            new_file_name = f"{store_name}：電力報告書{today_date_str}.xlsx"
            
            final_path = os.path.join(temp_dir, new_file_name)
            # shutil.copyではなく、openpyxl.save()がtemp_excel_pathに保存済みなので、名前を変更する
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
            st.warning("特に、CSVのヘッダー行が「年,月,日,時,...」の形式が崩れていないか、またE列以降に数値データが含まれているか確認してください。")
            st.exception(e)

if __name__ == "__main__":
    main_streamlit_app()
