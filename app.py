# app.py
import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import shutil
import chardet
import datetime
import openpyxl

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
            # まず全体をヘッダーなしで読み込み（バイナリから）
            df_full = pd.read_csv(io.BytesIO(raw), header=None, encoding=enc, keep_default_na=False)
            # ヘッダー行を探す（'年','月','日','時' を含む行）
            header_row_index = -1
            for i in range(df_full.shape[0]):
                row = df_full.iloc[i].astype(str).tolist()
                if all(x in row for x in ['年', '月', '日', '時']):
                    header_row_index = i
                    break
            if header_row_index == -1:
                continue

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
            # もし data 列数と cleaned_cols が合わない場合は調整
            if data.shape[1] != len(cleaned_cols):
                # 列数が違うときは不足する部分を埋める
                while len(cleaned_cols) < data.shape[1]:
                    cleaned_cols.append(f'Unnamed_{len(cleaned_cols)}')
                if len(cleaned_cols) > data.shape[1]:
                    cleaned_cols = cleaned_cols[:data.shape[1]]

            data.columns = cleaned_cols
            # 最低限 '年','月','日','時' が揃っているか確認
            if not all(col in data.columns for col in ['年', '月', '日', '時']):
                continue

            return data

        except Exception:
            continue

    raise Exception(f"CSVファイル '{uploaded_file.name}' を適切に読み込めませんでした（エンコーディング/形式を確認してください）。")

# ---------------------------
# Excel書き込み関数
# ---------------------------
def write_excel_reports(excel_path, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    """
    - df_before/df_after: '年','月','日','時','合計kWh','日付' を含むDataFrame
    - 0-23 時毎の平均を算出し、Sheet1 に C36-C59 (before), D36-D59 (after) として書き込む
    - まとめシートに期間・営業時間・店舗名を書き込む
    - 日別平均セル(C33/D33, まとめのB7/B8)は空欄にする
    """
    SHEET1 = "Sheet1"
    SUMMARY = "まとめ"

    try:
        wb = openpyxl.load_workbook(excel_path)
    except FileNotFoundError:
        st.error("Excelテンプレートが見つかりません。")
        return False

    # --- prepare metrics ---
    def hourly_mean_series(df):
        if df is None or df.empty:
            return pd.Series([0.0]*24, index=range(24), dtype=float)
        ser = df.groupby('時')['合計kWh'].mean()  # 時ごとの単純平均
        # index を int にして 0..23 に reindex（無ければ 0.0）
        ser.index = ser.index.astype(int)
        ser = ser.reindex(range(24), fill_value=0.0)
        return ser

    ser_before = hourly_mean_series(df_before)
    ser_after = hourly_mean_series(df_after)

    # --- Sheet1 書き込み ---
    if SHEET1 not in wb.sheetnames:
        wb.create_sheet(SHEET1)
    ws1 = wb[SHEET1]

    # C33/D33 は仕様どおり空欄（もしテンプレが式を期待しているなら上書きは避ける）
    try:
        ws1['C33'].value = None
        ws1['D33'].value = None
    except Exception:
        pass

    # C36 (row 36) ～ C59 (row 59) に 0時～23時を順に書き込む
    start_row = 36
    for hour in range(24):
        row = start_row + hour
        val_b = float(ser_before.loc[hour]) if not pd.isna(ser_before.loc[hour]) else 0.0
        val_a = float(ser_after.loc[hour]) if not pd.isna(ser_after.loc[hour]) else 0.0
        # 少数（例えば小数第3位以下）を整えて書きたい場合は round を使う
        ws1.cell(row=row, column=3, value=round(val_b, 4))  # C列
        ws1.cell(row=row, column=4, value=round(val_a, 4))  # D列

    # --- まとめシート 書き込み ---
    if SUMMARY not in wb.sheetnames:
        wb.create_sheet(SUMMARY)
    ws_sum = wb[SUMMARY]

    fmt = lambda d: f"{d.year}/{d.month}/{d.day}"
    ws_sum['H6'] = f"施工前：{fmt(start_before)}～{fmt(end_before)}（{(end_before - start_before).days + 1}日間）"
    ws_sum['H7'] = f"施工後(調光後)：{fmt(start_after)}～{fmt(end_after)}（{(end_after - start_after).days + 1}日間）"
    ws_sum['H8'] = operating_hours
    ws_sum['B1'] = f"{store_name}の使用電力比較報告書"
    # 日別平均セルは空にする
    try:
        ws_sum['B7'].value = None
        ws_sum['B8'].value = None
    except Exception:
        pass

    # 保存
    wb.save(excel_path)
    return True

# ---------------------------
# Streamlit アプリ本体
# ---------------------------
def main():
    st.set_page_config(layout="wide", page_title="電力データ自動処理アプリ")
    st.title("💡 電力データ自動処理アプリ")
    

    uploaded_csvs = st.file_uploader("📈 CSVデータ (複数可) をアップロードしてください", type=['csv'], accept_multiple_files=True)
    col1, col2 = st.columns(2)

    today = datetime.date.today()
    with col1:
        st.subheader("🗓️ 施工前")
        start_before = st.date_input("開始日 (施工前)", today - datetime.timedelta(days=30), key="start_b")
        end_before = st.date_input("終了日 (施工前)", today - datetime.timedelta(days=23), key="end_b")
    with col2:
        st.subheader("📅 施工後")
        start_after = st.date_input("開始日 (施工後)", today - datetime.timedelta(days=14), key="start_a")
        end_after = st.date_input("終了日 (施工後)", today - datetime.timedelta(days=7), key="end_a")

    operating_hours = st.text_input("営業時間", value="08:00-22:00")
    store_name = st.text_input("店舗名", value="大倉山店")

    st.markdown("---")

    if not uploaded_csvs:
        st.info("CSVファイルをアップロードすると実行ボタンが有効になります。")
        st.stop()

    if st.button("🚀 データ処理を実行して報告書を作成"):
        # 期間チェック（開始 <= 終了）
        if start_before > end_before or start_after > end_after:
            st.error("期間指定が不正です。開始日は終了日より前または同じ日にしてください。")
            st.stop()

        # テンプレ存在チェック
        if not os.path.exists(EXCEL_TEMPLATE_FILENAME):
            st.error(f"テンプレート '{EXCEL_TEMPLATE_FILENAME}' が見つかりません。アプリの実行フォルダに置いてください。")
            st.stop()

        # 一時フォルダ準備
        os.makedirs(TEMP_DIR, exist_ok=True)
        temp_excel_path = os.path.join(TEMP_DIR, EXCEL_TEMPLATE_FILENAME)
        shutil.copy(EXCEL_TEMPLATE_FILENAME, temp_excel_path)

        # --- CSV読み込み・統合 ---
        dfs = []
        try:
            for f in uploaded_csvs:
                df = detect_and_read_csv(f)
                dfs.append(df)
        except Exception as e:
            st.error("CSV読み込み時にエラーが発生しました。ファイルの形式やエンコーディングを確認してください。")
            st.exception(e)
            st.stop()

        if not dfs:
            st.error("CSVファイルからデータが読み取れませんでした。")
            st.stop()

        df_all = pd.concat(dfs, ignore_index=True)

        # 数値変換: 年, 月, 日, 時
        for col in ['年','月','日','時']:
            if col in df_all.columns:
                df_all[col] = pd.to_numeric(df_all[col], errors='coerce')
            else:
                st.error(f"CSVに必須カラム '{col}' が見つかりません。")
                st.stop()

        # 欠損行は除外
        df_all.dropna(subset=['年','月','日','時'], inplace=True)

        # 時の標準化: 1-24 の場合は -1 して 0-23 にする（1→0,24→23）
        if df_all['時'].max() > 23:
            df_all['時'] = df_all['時'].astype(int) - 1
            st.info("CSVの時刻が1-24形式だったため、0-23形式に変換しました。")

        df_all['時'] = df_all['時'].astype(int)

        # 消費カラムの特定（kWh_で始まるもの）
        consumption_cols = [c for c in df_all.columns if c.startswith('kWh_')]
        if not consumption_cols:
            st.error("E列以降に消費電力の数値カラムが見つかりません（kWh_で始まるカラム）。CSV形式を確認してください。")
            st.stop()

        # 数値変換（NaNを0に）
        for c in consumption_cols:
            df_all[c] = pd.to_numeric(df_all[c], errors='coerce').fillna(0.0)

        # 同じ (年,月,日,時) を合算（行同士の合算）
        grouped = df_all.groupby(['年','月','日','時'], as_index=False)[consumption_cols].sum()
        # 合算結果を一列にまとめる
        grouped['合計kWh'] = grouped[consumption_cols].sum(axis=1)

        # 日付列を作成
        grouped['日付'] = pd.to_datetime(
            grouped['年'].astype(int).astype(str) + "-" +
            grouped['月'].astype(int).astype(str) + "-" +
            grouped['日'].astype(int).astype(str),
            format='%Y-%m-%d', errors='coerce'
        ).dt.date
        grouped.dropna(subset=['日付'], inplace=True)

        # 期間フィルタ
        df_before = grouped[(grouped['日付'] >= start_before) & (grouped['日付'] <= end_before)].copy()
        df_after = grouped[(grouped['日付'] >= start_after) & (grouped['日付'] <= end_after)].copy()

        # 欠損チェック（期待値との比較）
        days_b = (end_before - start_before).days + 1
        expected_b = days_b * 24
        found_b = df_before.shape[0]
        if df_before.empty or found_b < expected_b * 0.95:
            st.warning(f"施工前期間の読み取り件数が少ない可能性: 期待 {expected_b} 件 / 実際 {found_b} 件")

        days_a = (end_after - start_after).days + 1
        expected_a = days_a * 24
        found_a = df_after.shape[0]
        if df_after.empty or found_a < expected_a * 0.95:
            st.warning(f"施工後期間の読み取り件数が少ない可能性: 期待 {expected_a} 件 / 実際 {found_a} 件")

        # Excel書き込み
        success = write_excel_reports(temp_excel_path, df_before, df_after,
                                      start_before, end_before, start_after, end_after,
                                      operating_hours, store_name)
        if not success:
            st.error("Excelへの書き込みに失敗しました。")
            st.stop()

        # 保存ファイル名リネームとダウンロード提供
        today_str = datetime.date.today().strftime('%Y%m%d')
        out_name = f"{store_name}_電力報告書_{today_str}.xlsx"
        final_path = os.path.join(TEMP_DIR, out_name)
        try:
            os.replace(temp_excel_path, final_path)
        except Exception:
            shutil.copy(temp_excel_path, final_path)

        with open(final_path, "rb") as f:
            st.success("✅ 処理完了しました。以下からダウンロードしてください。")
            st.download_button(
                label="⬇️ 報告書をダウンロード",
                data=f,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
