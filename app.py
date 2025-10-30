import streamlit as st
import pandas as pd
import os
import glob
import sys
import chardet
import openpyxl
from openpyxl.utils import cell
import datetime

# ======================================================
# 💡 注意: 以下の関数は、元のスクリプトの関数をそのまま使用しています
# ======================================================

# --- CSV読み込み関数 (変更なし) ---
def detect_and_read_csv(file_path):
    # ... (元の detect_and_read_csv 関数のロジックをここにペースト) ...
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    
    detected_encoding = chardet.detect(raw_data)['encoding']
    encodings_to_try = ['cp932', 'shift_jis', 'utf-8']
    
    if detected_encoding and detected_encoding.lower() not in encodings_to_try:
        encodings_to_try.append(detected_encoding.lower())

    for encoding in encodings_to_try:
        try:
            # pandas.read_csvのfilepath_or_bufferにファイルパスではなく、ファイルオブジェクトを渡す
            df = pd.read_csv(file_path, header=1, encoding=encoding)
            
            if '年' in df.columns:
                 # Streamlit環境では、標準出力はログに出るのみ
                 return df
            else:
                 continue

        except Exception:
            continue
            
    raise UnicodeDecodeError(f"ファイルは、一般的な日本語エンコーディングで読み込めませんでした。")


# --- Excel書き込み関数 (Streamlit用に引数を微調整) ---
def write_excel_reports(excel_file, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    # ... (元の write_excel_reports 関数のロジックをここにペースト) ...
    SHEET1_NAME = 'Sheet1'
    SUMMARY_SHEET_NAME = 'まとめ'
    
    try:
        workbook = openpyxl.load_workbook(excel_file)
    except FileNotFoundError:
        st.error(f"エラー: Excelファイル '{excel_file}' が見つかりません。")
        return

    # Sheet1: 24時間別平均の書き込み
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

    # まとめシート: 期間 (H6, H7), 営業時間 (H8), タイトル (B1) の書き込み
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
    
    workbook.save(excel_file)
    
    return workbook

# ======================================================
# 💡 Streamlitメインアプリケーション
# ======================================================

def main_streamlit_app():
    st.set_page_config(layout="wide", page_title="電力データ報告書作成アプリ")
    st.title("💡 電力データ自動処理アプリ")
    st.markdown("### Step 1: ファイルのアップロード")
    
    # --- 1. ファイルのアップロード ---
    col_up1, col_up2 = st.columns(2)
    
    # CSVファイルのアップロード
    uploaded_csvs = col_up1.file_uploader(
        "📈 CSVデータ (複数可) をアップロードしてください",
        type=['csv'],
        accept_multiple_files=True
    )
    
    # Excelテンプレートのアップロード
    uploaded_excel = col_up2.file_uploader(
        "📄 Excelテンプレートファイルをアップロードしてください",
        type=['xlsx'],
        accept_multiple_files=False
    )
    
    if uploaded_csvs and uploaded_excel:
        st.success(f"CSVファイル {len(uploaded_csvs)}個 と Excelファイル 1個 が準備できました。")
        st.markdown("---")
        st.markdown("### Step 2: 期間と情報の入力")
    else:
        st.warning("処理を開始するには、CSVデータとExcelテンプレートの両方をアップロードしてください。")
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
        # 実行ロジックを try/except でラップ
        try:
            # テンポラリフォルダのセットアップ
            # Streamlit環境ではメモリ上で処理するのが基本だが、openpyxl/pandas連携のためディスクに一時保存
            temp_dir = "temp_data"
            os.makedirs(temp_dir, exist_ok=True)
            
            # --- a) アップロードされたファイルのテンポラリ保存 ---
            # Excelファイルを保存
            excel_path = os.path.join(temp_dir, uploaded_excel.name)
            with open(excel_path, "wb") as f:
                f.write(uploaded_excel.getbuffer())
                
            # CSVファイルを保存し、リストを作成
            csv_paths = []
            for csv_file in uploaded_csvs:
                csv_path = os.path.join(temp_dir, csv_file.name)
                with open(csv_path, "wb") as f:
                    f.write(csv_file.getbuffer())
                csv_paths.append(csv_path)

            # --- b) データ統合と前処理 (元のロジックを呼び出し) ---
            
            # データの読み込みと統合
            all_data = []
            for csv_path in csv_paths:
                df = detect_and_read_csv(csv_path) # 修正された CSV リーダーを使用
                all_data.append(df)
            df_combined = pd.concat(all_data, ignore_index=True)
            
            # データ前処理（元のロジックをここに貼り付ける）
            # ... (前処理ロジックの簡略化 - 実際は元の完全なロジックが必要です)
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
            
            # --- d) Excel書き込みとファイル名変更 ---
            
            # Pandasでのデータシート書き込み
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                writer.sheets = {ws.title: ws for ws in openpyxl.load_workbook(excel_path).worksheets}
                df_combined.to_excel(writer, sheet_name='元データ', index=False, if_sheet_exists='replace') 
                df_before_full.to_excel(writer, sheet_name='施工前', index=False, if_sheet_exists='replace')   
                df_after_full.to_excel(writer, sheet_name='施工後（調光後）', index=False, if_sheet_exists='replace')

            # openpyxlでSheet1とまとめシートを更新
            write_excel_reports(excel_path, df_before, df_after, start_b, end_b, start_a, end_a, operating_hours, store_name)
            
            
            # --- e) ファイル名の変更とダウンロードの準備 ---
            
            today_date_str = datetime.date.today().strftime('%Y%m%d')
            new_file_name = f"{store_name}：電力報告書{today_date_str}.xlsx"
            
            # ファイル名を変更してダウンロード用のパスを取得
            final_path = os.path.join(temp_dir, new_file_name)
            os.rename(excel_path, final_path)
            
            # ファイルを読み込み、ダウンロードボタンを作成
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
            st.exception(e)

if __name__ == "__main__":
    main_streamlit_app()
