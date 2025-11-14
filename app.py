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
import matplotlib.pyplot as plt

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

            # カラム名整形：A-D はそのまま、E以降は kWh_1...
            cleaned_cols = []
            k = 1
            for i, col in enumerate(header):
                if i < 4:
                    cleaned_cols.append(str(col))
                else:
                    cleaned_cols.append(f'kWh_{k}')
                    k += 1

            # 読み込んだ行数とヘッダー長がずれる場合の補正
            if data.shape[1] != len(cleaned_cols):
                # 足りないなら Unnamed を追加、余るなら切る
                while len(cleaned_cols) < data.shape[1]:
                    cleaned_cols.append(f'Unnamed_{len(cleaned_cols)}')
                if len(cleaned_cols) > data.shape[1]:
                    cleaned_cols = cleaned_cols[:data.shape[1]]

            data.columns = cleaned_cols

            # 必須カラムチェック
            if not all(col in data.columns for col in ['年', '月', '日', '時']):
                continue

            return data

        except Exception:
            continue

    raise Exception(f"CSVファイル '{getattr(uploaded_file, 'name', 'unknown')}' を適切に読み込めませんでした（エンコーディング/形式を確認してください）。")

# ---------------------------
# Excel書き込み関数
# ---------------------------
def write_excel_reports(excel_path, df_before, df_after, start_before, end_before, start_after, end_after, operating_hours, store_name):
    """
    - df_before/df_after: '年','月','日','時','合計kWh','日付' を含むDataFrame
    - 0-23 時毎の平均を算出し、Sheet1 に C36-C59 (before), D36-D59 (after) として書き込む
    - まとめシートに期間・営業時間・店舗名を書き込む
    """
    SHEET1 = "Sheet1"
    SUMMARY = "まとめ"

    try:
        wb = openpyxl.load_workbook(excel_path)
    except FileNotFoundError:
        st.error("Excelテンプレートが見つかりません。")
        return False

    def hourly_mean_series(df):
        if df is None or df.empty:
            return pd.Series([0.0]*24, index=range(24), dtype=float)
        ser = df.groupby('時')['合計kWh'].mean()
        ser.index = ser.index.astype(int)
        ser = ser.reindex(range(24), fill_value=0.0)
        return ser

    ser_before = hourly_mean_series(df_before)
    ser_after = hourly_mean_series(df_after)

    # Sheet1 書き込み
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

    # まとめシート 書き込み
    if SUMMARY not in wb.sheetnames:
        wb.create_sheet(SUMMARY)
    ws_sum = wb[SUMMARY]

    fmt = lambda d: f"{d.year}/{d.month}/{d.day}"
    ws_sum['H6'] = f"施工前：{fmt(start_before)}～{fmt(end_before)}（{(end_before - start_before).days + 1}日間）"
    ws_sum['H7'] = f"施工後(調光後)：{fmt(start_after)}～{fmt(end_after)}（{(end_after - start_after).days + 1}日間）"
    ws_sum['H8'] = operating_hours
    ws_sum['B1'] = f"{store_name}の使用電力比較報告書"


    wb.save(excel_path)
    return True

# ---------------------------
# ヘルパー: 集計 → 時間平均・差分テーブル作成
# ---------------------------
def build_hourly_comparison(df_before, df_after):
    """
    df_before/after: grouped dataframe with columns '年','月','日','時','合計kWh','日付'
    returns a DataFrame with index 0..23 and columns:
    before_avg, after_avg, savings (before-after), savings_pct
    """
    def hourly_mean(df):
        if df is None or df.empty:
            return pd.Series([0.0]*24, index=range(24), dtype=float)
        s = df.groupby('時')['合計kWh'].mean()
        s.index = s.index.astype(int)
        s = s.reindex(range(24), fill_value=0.0)
        return s

    b = hourly_mean(df_before)
    a = hourly_mean(df_after)

    df = pd.DataFrame({
        'hour': range(24),
        'before_avg_kWh': [float(b.loc[h]) for h in range(24)],
        'after_avg_kWh': [float(a.loc[h]) for h in range(24)]
    })
    df['savings_kWh'] = df['before_avg_kWh'] - df['after_avg_kWh']
    # %節電（beforeが0のときは None）
    df['savings_pct'] = df.apply(lambda r: (r['savings_kWh'] / r['before_avg_kWh'] * 100) if r['before_avg_kWh'] != 0 else None, axis=1)
    return df

# ---------------------------
# Streamlit アプリ本体
# ---------------------------
def main():
    st.set_page_config(layout="wide", page_title="電力データ自動処理アプリ")
    st.title("💡 電力データ自動処理アプリ（施工前/施工後 比較）")
    st.markdown("CSVをアップロードして、施工前/施工後の0-23時ごとの平均を計算し、Excelテンプレに出力します。")

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
        # 期間チェック
        if start_before > end_before or start_after > end_after:
            st.error("期間指定が不正です。開始日は終了日より前または同じ日にしてください。")
            st.stop()

        # テンプレ存在チェック
        if not os.path.exists(EXCEL_TEMPLATE_FILENAME):
            st.error(f"テンプレート '{EXCEL_TEMPLATE_FILENAME}' が見つかりません。アプリの実行フォルダに置いてください。")
            st.stop()

        os.makedirs(TEMP_DIR, exist_ok=True)
        temp_excel_path = os.path.join(TEMP_DIR, EXCEL_TEMPLATE_FILENAME)
        shutil.copy(EXCEL_TEMPLATE_FILENAME, temp_excel_path)

        # --- CSV読み込み ---
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

        # 必須カラムを数値化（失敗はNaN）
        for col in ['年','月','日','時']:
            if col in df_all.columns:
                df_all[col] = pd.to_numeric(df_all[col], errors='coerce')
            else:
                st.error(f"CSVに必須カラム '{col}' が見つかりません。")
                st.stop()

        # === 不正行排除ロジック ===
        # (1) 年/月/日/時 がいずれか欠けている行は除外
        df_all = df_all.dropna(subset=['年','月','日','時'])

        # (2) 時の範囲（0~24）のみ残す（まず広く許容）
        df_all = df_all[df_all['時'].between(0, 24)]

        # (3) もし1-24表記だったら 0-23 に変換
        if df_all['時'].max() > 23:
            # 整数化して -1
            df_all['時'] = df_all['時'].astype(int) - 1
            st.info("CSVの時刻が1-24形式だったため、0-23形式に変換しました。")

        # (4) 最終チェック：0-23 のみ残す
        df_all = df_all[df_all['時'].between(0, 23)]

        # (5) 年/月/日/時 を整数化（例: 2024.0 -> 2024）
        df_all[['年','月','日','時']] = df_all[['年','月','日','時']].astype(int)

        # 消費カラム
        consumption_cols = [c for c in df_all.columns if c.startswith('kWh_')]
        if not consumption_cols:
            st.error("E列以降に消費電力の数値カラムが見つかりません（kWh_で始まるカラム）。CSV形式を確認してください。")
            st.stop()

        # 数値変換（NaNは0.0に）
        for c in consumption_cols:
            df_all[c] = pd.to_numeric(df_all[c], errors='coerce').fillna(0.0)

        # --- E列以降を合算して '合計kWh' を作る（行ごと）
        df_all['合計kWh'] = df_all[consumption_cols].sum(axis=1)

        # --- 同一 (年,月,日,時) を合算（複数行がある場合にまとめる）
        grouped = df_all.groupby(['年','月','日','時'], as_index=False)[consumption_cols + ['合計kWh']].sum()
        # 合算後の合計kWh は再計算（安全のため）
        grouped['合計kWh'] = grouped[consumption_cols].sum(axis=1) if consumption_cols else grouped['合計kWh']

        # 日付列を作る
        grouped['日付'] = pd.to_datetime(
            grouped['年'].astype(int).astype(str) + "-" +
            grouped['月'].astype(int).astype(str) + "-" +
            grouped['日'].astype(int).astype(str),
            format='%Y-%m-%d', errors='coerce'
        ).dt.date
        grouped = grouped.dropna(subset=['日付'])

        # 期間で分割
        df_before = grouped[(grouped['日付'] >= start_before) & (grouped['日付'] <= end_before)].copy()
        df_after = grouped[(grouped['日付'] >= start_after) & (grouped['日付'] <= end_after)].copy()

        # 欠損チェック
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

        # --- 集計・比較テーブル作成 ---
        comp = build_hourly_comparison(df_before, df_after)
        comp_display = comp.copy()
        comp_display['savings_pct'] = comp_display['savings_pct'].apply(lambda x: f"{x:.1f}%" if x is not None else "-")
        comp_display['before_avg_kWh'] = comp_display['before_avg_kWh'].round(4)
        comp_display['after_avg_kWh'] = comp_display['after_avg_kWh'].round(4)
        comp_display['savings_kWh'] = comp_display['savings_kWh'].round(4)

        st.subheader("時間帯別平均（0～23時）")
        st.dataframe(comp_display.rename(columns={
            'hour': '時刻',
            'before_avg_kWh': '施工前 平均(kWh)',
            'after_avg_kWh': '施工後 平均(kWh)',
            'savings_kWh': '差分(kWh)',
            'savings_pct': '差分(%)'
        }), use_container_width=True)

        # 全体の合計節電量（平均値の合算ではなく、時間帯別平均の差分を24時間合算）
        total_savings_kWh = comp['savings_kWh'].sum()
        # 全体節電率（中央値的ではなく、合計比率）： (sum(before_avg) - sum(after_avg))/sum(before_avg)
        sum_before = comp['before_avg_kWh'].sum()
        sum_after = comp['after_avg_kWh'].sum()
        total_savings_pct = (sum_before - sum_after) / sum_before * 100 if sum_before != 0 else None

        st.markdown("---")
        col_a, col_b, col_c = st.columns([1,1,1])
        col_a.metric("合計：施工前平均 (24h合計)", f"{sum_before:.4f} kWh")
        col_b.metric("合計：施工後平均 (24h合計)", f"{sum_after:.4f} kWh")
        col_c.metric("合計節電量 (24h)", f"{total_savings_kWh:.4f} kWh", f"{total_savings_pct:.1f}% " if total_savings_pct is not None else "")

       

        # --- Excel書き込み ---
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
