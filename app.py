import streamlit as st
import pandas as pd
import io
import chardet
import datetime
import openpyxl


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
            df_full = pd.read_csv(io.BytesIO(raw), header=None, encoding=enc, keep_default_na=False)
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

            cleaned_cols = []
            k = 1
            for i, col in enumerate(header):
                if i < 4:
                    cleaned_cols.append(str(col))
                else:
                    cleaned_cols.append(f'kWh_{k}')
                    k += 1

            if data.shape[1] != len(cleaned_cols):
                while len(cleaned_cols) < data.shape[1]:
                    cleaned_cols.append(f'Unnamed_{len(cleaned_cols)}')
                if len(cleaned_cols) > data.shape[1]:
                    cleaned_cols = cleaned_cols[:data.shape[1]]

            data.columns = cleaned_cols

            if not all(col in data.columns for col in ['年', '月', '日', '時']):
                continue

            return data

        except Exception:
            continue

    raise Exception(f"CSVファイル '{getattr(uploaded_file, 'name', 'unknown')}' を適切に読み込めませんでした（エンコーディング/形式を確認してください）。")

# ---------------------------
# Excel書き込み関数
# ---------------------------
def write_excel_reports(template_bytes, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    SHEET1 = "Sheet1"
    SUMMARY = "まとめ"

    try:
        wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    except Exception as e:
        print(f"Error loading template: {e}")
        return None

    def hourly_mean_series(df):
        if df is None or df.empty:
            return pd.Series([0.0]*24, index=range(24), dtype=float)
        
        # ファイル(計測箇所)ごとに平均を計算し、合算する
        mean_series = pd.Series([0.0]*24, index=range(24), dtype=float)
        for fid in df['file_id'].unique():
            df_file = df[df['file_id'] == fid]
            ser = df_file.groupby('時')['合計kWh'].mean()
            ser.index = ser.index.astype(int)
            ser = ser.reindex(range(24), fill_value=0.0)
            mean_series += ser
        return mean_series

    ser_before = hourly_mean_series(df_before)
    ser_after = hourly_mean_series(df_after)

    if SHEET1 not in wb.sheetnames:
        wb.create_sheet(SHEET1)
    ws1 = wb[SHEET1]

    start_row = 36
    for hour in range(24):
        row = start_row + hour
        val_b = float(ser_before.loc[hour]) if not pd.isna(ser_before.loc[hour]) else 0.0
        val_a = float(ser_after.loc[hour]) if not pd.isna(ser_after.loc[hour]) else 0.0
        ws1.cell(row=row, column=3, value=round(val_b, 4))
        ws1.cell(row=row, column=4, value=round(val_a, 4))

    if SUMMARY not in wb.sheetnames:
        wb.create_sheet(SUMMARY)
    ws_sum = wb[SUMMARY]

    fmt = lambda d: f"{d.year}/{d.month}/{d.day}"
    ws_sum['H6'] = f"施工前：{fmt(start_before)}～{fmt(end_before)}（{(end_before - start_before).days + 1}日間）"
    ws_sum['H7'] = f"施工後(調光後)：{fmt(start_after)}～{fmt(end_after)}（{(end_after - start_after).days + 1}日間）"
    ws_sum['H8'] = operating_hours
    ws_sum['B1'] = f"{store_name}の使用電力比較報告書"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ---------------------------
# Streamlit アプリ本体
# ---------------------------
def main():
    st.set_page_config(layout="wide", page_title="電力データ自動処理アプリ")
    st.title("💡 電力データ自動処理アプリ（施工前/施工後 比較）")
    st.markdown("CSVとExcelテンプレートをアップロードして、データ処理と報告書作成を行います。")

    template_file = st.file_uploader(
        "📥 Excelテンプレートファイル (電力報告テンプレート.xlsxなど) をアップロードしてください", 
        type=['xlsx']
    )
    uploaded_csvs = st.file_uploader("📈 CSVデータ (複数可) をアップロードしてください", type=['csv'], accept_multiple_files=True)
    
    st.markdown("---")

    col1, col2 = st.columns(2)
    today = datetime.date.today()

    with col1:
        st.subheader("🗓️ 施工前期間")
        start_before = st.date_input("開始日 (施工前)", today - datetime.timedelta(days=30), key="start_b")
        end_before = st.date_input("終了日 (施工前)", today - datetime.timedelta(days=23), key="end_b")
    with col2:
        st.subheader("📅 施工後期間")
        start_after = st.date_input("開始日 (施工後)", today - datetime.timedelta(days=14), key="start_a")
        end_after = st.date_input("終了日 (施工後)", today - datetime.timedelta(days=7), key="end_a")

    operating_hours = st.text_input("営業時間", value="08:00-22:00")
    store_name = st.text_input("店舗名", value="大倉山店")

    st.markdown("---")

    if not uploaded_csvs or not template_file:
        st.info("CSVファイルとExcelテンプレートファイルをアップロードすると実行ボタンが有効になります。")
        st.stop()

    if st.button("🚀 データ処理を実行して報告書を作成"):
        if start_before > end_before or start_after > end_after:
            st.error("期間指定が不正です。開始日は終了日より前または同じ日にしてください。")
            st.stop()
        
        template_bytes = template_file.read()

        # --- CSV読み込み ---
        dfs = []
        try:
            for i, f in enumerate(uploaded_csvs):
                df = detect_and_read_csv(f)
                df['file_id'] = i  # ファイル識別用のIDを追加
                dfs.append(df)
        except Exception as e:
            st.error("CSV読み込み時にエラーが発生しました。ファイルの形式やエンコーディングを確認してください。")
            st.exception(e)
            st.stop()

        if not dfs:
            st.error("CSVファイルからデータが読み取れませんでした。")
            st.stop()

        df_all = pd.concat(dfs, ignore_index=True)

        for col in ['年','月','日','時']:
            if col in df_all.columns:
                df_all[col] = pd.to_numeric(df_all[col], errors='coerce')
            else:
                st.error(f"CSVに必須カラム '{col}' が見つかりません。")
                st.stop()

        df_all = df_all.dropna(subset=['年','月','日','時'])
        df_all = df_all[df_all['時'].between(0, 24)]

        if df_all['時'].max() > 23:
            df_all['時'] = df_all['時'].astype(int) - 1
            st.info("CSVの時刻が1-24形式だったため、0-23形式に変換しました。")

        df_all = df_all[df_all['時'].between(0, 23)]
        df_all[['年','月','日','時']] = df_all[['年','月','日','時']].astype(int)

        consumption_cols = [c for c in df_all.columns if c.startswith('kWh_')]
        if not consumption_cols:
            st.error("E列以降に消費電力の数値カラムが見つかりません（kWh_で始まるカラム）。CSV形式を確認してください。")
            st.stop()

        for c in consumption_cols:
            df_all[c] = pd.to_numeric(df_all[c], errors='coerce').fillna(0.0)

        df_all['合計kWh'] = df_all[consumption_cols].sum(axis=1)

        # file_id も含めてgroupbyする（同一ファイル内の同一日時をまとめる）
        grouped = df_all.groupby(['file_id', '年','月','日','時'], as_index=False)[consumption_cols + ['合計kWh']].sum()
        grouped['合計kWh'] = grouped[consumption_cols].sum(axis=1) if consumption_cols else grouped['合計kWh']

        grouped['日付'] = pd.to_datetime(
            grouped['年'].astype(int).astype(str) + "-" +
            grouped['月'].astype(int).astype(str) + "-" +
            grouped['日'].astype(int).astype(str),
            format='%Y-%m-%d', errors='coerce'
        ).dt.date
        grouped = grouped.dropna(subset=['日付'])

        df_before = grouped[(grouped['日付'] >= start_before) & (grouped['日付'] <= end_before)].copy()
        df_after = grouped[(grouped['日付'] >= start_after) & (grouped['日付'] <= end_after)].copy()

        days_b = (end_before - start_before).days + 1
        expected_b = days_b * 24
        found_b = df_before.shape[0] if not df_before.empty else 0
        
        days_a = (end_after - start_after).days + 1
        expected_a = days_a * 24
        found_a = df_after.shape[0] if not df_after.empty else 0
        
        # --- Excel書き込み（メモリ上で完結） ---
        excel_buffer = write_excel_reports(
            template_bytes, df_before, df_after,
            start_before, end_before, start_after, end_after,
            operating_hours, store_name
        )
        
        if excel_buffer is None:
            st.error("Excelテンプレートの読み込みまたはデータ書き込みに失敗しました。")
            st.stop()

        today_str = datetime.date.today().strftime('%Y%m%d')
        out_name = f"{store_name}_電力報告書_{today_str}.xlsx"
        
        st.success("✅ 処理完了しました。以下からダウンロードしてください。")
        st.download_button(
            label="⬇️ 報告書をダウンロード",
            data=excel_buffer,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
