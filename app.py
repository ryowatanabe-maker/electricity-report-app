import pandas as pd
import os
import glob
import sys
import chardet
import openpyxl
from openpyxl.utils import cell
import datetime

# chardet, openpyxl, pandasがインストールされていることを前提とします。

def detect_and_read_csv(file_path):
    """ファイルの内容を読み込み、日本語CSVに最適化された順序でエンコーディングを試行してDataFrameを返す"""
    
    with open(file_path, 'rb') as f:
        raw_data = f.read()
    
    detected_encoding = chardet.detect(raw_data)['encoding']
    encodings_to_try = ['cp932', 'shift_jis', 'utf-8']
    
    if detected_encoding and detected_encoding.lower() not in encodings_to_try:
        encodings_to_try.append(detected_encoding.lower())

    for encoding in encodings_to_try:
        try:
            df = pd.read_csv(file_path, header=1, encoding=encoding)
            
            if '年' in df.columns:
                 print(f"    - '{encoding}' で日本語ヘッダーを正常に読み込みました。")
                 return df
            else:
                 continue

        except Exception:
            continue
            
    raise UnicodeDecodeError(f"ファイル '{file_path}' は、一般的な日本語エンコーディングで読み込めませんでした。")

def write_excel_reports(excel_file, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    """
    Sheet1に24時間分の時間帯別の平均値と、まとめシートに測定期間、営業時間、店舗名を書き込む。
    """
    SHEET1_NAME = 'Sheet1'
    SUMMARY_SHEET_NAME = 'まとめ'
    
    try:
        workbook = openpyxl.load_workbook(excel_file)
    except FileNotFoundError:
        print(f"エラー: Excelファイル '{excel_file}' が見つかりません。")
        return

    # --- 1. Sheet1: 24時間別平均の書き込み (C36～D59) ---
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
    print(f"    - 【{SHEET1_NAME}】の C36～D59 に 01:00～24:00 の時間帯別平均を反映しました。")

    # --- 2. まとめシート: 期間 (H6, H7), 営業時間 (H8), タイトル (B1) の書き込み ---
    if SUMMARY_SHEET_NAME not in workbook.sheetnames:
        workbook.create_sheet(SUMMARY_SHEET_NAME)
    ws_summary = workbook[SUMMARY_SHEET_NAME]

    # 期間情報の計算と書き込み (H6, H7)
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
    
    # H8に営業時間を書き込み
    ws_summary['H8'] = operating_hours
    
    # B1に店舗名を含んだタイトルを書き込み
    ws_summary['B1'] = f"{store_name}の使用電力比較報告書"
    
    workbook.save(excel_file)
    print(f"    - 【{SUMMARY_SHEET_NAME}】シートの H6/H7/H8 および B1 に情報を反映しました。")


def split_data_and_calculate_average_automatic():
    """
    メイン処理: ファイル検出、データ統合、分割、計算、Excel書き込み、ファイル名変更
    """
    
    try:
        import chardet
        import openpyxl
    except ImportError:
        print("エラー: 必要なライブラリが見つかりません。pip install chardet openpyxl を実行してください。")
        sys.exit()

    # --- 1. ファイルの自動検出 ---
    excel_files = sorted(glob.glob('*.xlsx'))
    if len(excel_files) == 0:
        print("エラー: Excelファイル（*.xlsx）が見つかりませんでした。")
        sys.exit()
        
    PREFERRED_EXCEL_FILE = '富士川店：電力報告250130.xlsx'
    EXCEL_FILE = PREFERRED_EXCEL_FILE if PREFERRED_EXCEL_FILE in excel_files else excel_files[0]
    
    print(f"✅ 書き込み対象 Excelファイル: '{EXCEL_FILE}'")

    csv_files = glob.glob('*.csv')
    if not csv_files:
        print("エラー: 実行フォルダ内にCSVファイル（*.csv）が見つかりませんでした。")
        sys.exit()

    print(f"✅ 以下のCSVファイルを検出しました: {csv_files}")
    
    # --- 2. 複数CSVファイルの統合 ---
    all_data = []
    for csv_file in csv_files:
        try:
            df = detect_and_read_csv(csv_file)
            all_data.append(df)
        except UnicodeDecodeError as e:
            print(f"エラー: '{csv_file}' の読み込み中に文字コードエラーが発生しました: {e}")
            sys.exit()
    
    df_combined = pd.concat(all_data, ignore_index=True)
    
    # --- 3. データ前処理 ---
    df_combined['年'] = pd.to_numeric(df_combined['年'], errors='coerce').astype('Int64')
    df_combined['月'] = pd.to_numeric(df_combined['月'], errors='coerce').astype('Int64')
    df_combined['日'] = pd.to_numeric(df_combined['日'], errors='coerce').astype('Int64')
    df_combined.dropna(subset=['年', '月', '日'], inplace=True)
    df_combined['日付'] = pd.to_datetime(
        df_combined['年'].astype(str) + '-' + df_combined['月'].astype(str) + '-' + df_combined['日'].astype(str), 
        format='%Y-%m-%d', errors='coerce'
    ).dt.date
    df_combined.dropna(subset=['日付'], inplace=True)
    df_combined = df_combined.dropna(axis=1, how='all')
    
    datetime_cols = ['年', '月', '日', '時', '日付']
    consumption_cols = [col for col in df_combined.columns if col not in datetime_cols and not col.startswith('Unnamed:')]
    
    if not consumption_cols:
        print("エラー: E列以降に消費電力データ（kWhや回路データ）のカラムが見つかりませんでした。")
        sys.exit()

    for col in consumption_cols:
        df_combined[col] = pd.to_numeric(df_combined[col], errors='coerce').fillna(0)
    df_combined['合計kWh'] = df_combined[consumption_cols].sum(axis=1)
    
    print("\n全データが正常に統合されました。\n---")
    
    # --- 4. ユーザー入力とデータ分割 ---
    
    start_before_str = input("【施工前】期間の開始日を入力（例: 2025-09-21）: ")
    end_before_str = input("【施工前】期間の終了日を入力（例: 2025-09-26）: ")
    start_after_str = input("【施工後】期間の開始日を入力（例: 2025-10-10）: ")
    end_after_str = input("【施工後】期間の終了日を入力（例: 2025-10-15）: ")

    try:
        start_before = datetime.datetime.strptime(start_before_str, '%Y-%m-%d').date()
        end_before = datetime.datetime.strptime(end_before_str, '%Y-%m-%d').date()
        start_after = datetime.datetime.strptime(start_after_str, '%Y-%m-%d').date()
        end_after = datetime.datetime.strptime(end_after_str, '%Y-%m-%d').date()
    except ValueError:
        print("エラー: 日付の形式が正しくありません。'YYYY-MM-DD'形式で入力してください。")
        sys.exit()
    
    operating_hours_input = input("営業時間を入力してください（例: 8:00-22:00）: ")
    store_name_input = input("店舗名を入力してください（例: 大倉山店）: ")
    
    # 日付でデータを抽出
    df_before_full = df_combined[(df_combined['日付'] >= start_before) & (df_combined['日付'] <= end_before)].copy()
    df_after_full = df_combined[(df_combined['日付'] >= start_after) & (df_combined['日付'] <= end_after)].copy()
    
    df_before = df_before_full.copy()
    df_after = df_after_full.copy()
    
    # --- 5. Excel書き込みと分析 ---
    
    SOURCE_SHEET = '元データ'
    BEFORE_SHEET = '施工前'
    AFTER_SHEET = '施工後（調光後）' 
    
    # 1. Pandasでデータシートを書き込み
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_combined.to_excel(writer, sheet_name=SOURCE_SHEET, index=False) 
            df_before_full.to_excel(writer, sheet_name=BEFORE_SHEET, index=False)   
            df_after_full.to_excel(writer, sheet_name=AFTER_SHEET, index=False)
    except Exception as e:
        print(f"Pandasでのデータシート書き込み中にエラーが発生しました: {e}")
        sys.exit()

    # 2. openpyxlでSheet1の時間帯別平均と、まとめシートの期間/営業時間/店舗名を書き込み
    try:
        write_excel_reports(EXCEL_FILE, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours_input, store_name_input)
        
        # --- ファイル名変更処理 ---
        # 1. 現在の日付を取得し、YYYYMMDD形式にフォーマット
        today_date_str = datetime.date.today().strftime('%Y%m%d')
        
        # 2. 新しいファイル名を構築
        new_file_name = f"{store_name_input}：電力報告書{today_date_str}.xlsx"
        
        # 3. ファイル名が異なる場合のみリネームを実行
        if EXCEL_FILE != new_file_name:
            if os.path.exists(new_file_name):
                 print(f"⚠️ 警告: 新しいファイル名 '{new_file_name}' が既に存在します。リネームをスキップします。")
            else:
                 os.rename(EXCEL_FILE, new_file_name)
                 print(f"✅ Excelファイル名を '{EXCEL_FILE}' から '{new_file_name}' に変更しました。")
                 EXCEL_FILE = new_file_name # 以降の表示のために更新

        print("\n---")
        print("✅ 全ての処理が完了しました。")
        print(f"【Sheet1】と【まとめ】シートに結果を反映しました。")
        
        # --- 結果の表示 ---
        avg_before_total = df_before['合計kWh'].mean()
        avg_after_total = df_after['合計kWh'].mean()
        print(f"\nExcelファイル名: {EXCEL_FILE}")
        print("=== 24時間期間の平均消費電力 (kWh/h) ===")
        print(f"【施工前】期間の平均: {avg_before_total:.3f} kWh/h")
        print(f"【施工後】期間の平均: {avg_after_total:.3f} kWh/h")

    except Exception as e:
        print(f"Excelへの書き込みまたはファイル名変更中にエラーが発生しました: {e}")

# --- 関数を実行 ---
if __name__ == "__main__":
    split_data_and_calculate_average_automatic()