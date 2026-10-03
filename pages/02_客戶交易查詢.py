import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="客戶交易與明細深度查詢工具", layout="wide", page_icon="🔍")

st.title("🔍 客戶交易與品項明細深度查詢系統")
st.write("💡 上傳銷貨明細檔（支援 Excel/CSV、多檔合併），即可依照**客戶編號**、**日期區間**、**交易單號**、**有無載具**等條件精準查詢總金額與品項明細。")

def clean_customer_id(x):
    if pd.isnull(x):
        return ""
    s = str(x).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def try_parse_date(date_val):
    if pd.isnull(date_val):
        return None
    if isinstance(date_val, (datetime, pd.Timestamp)):
        return date_val.strftime('%Y-%m-%d')
    s = str(date_val).strip().split(' ')[0]
    m1 = re.match(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m1:
        return f"{m1.group(1)}-{int(m1.group(2)):02d}-{int(m1.group(3)):02d}"
    m2 = re.match(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', s)
    if m2:
        return f"{m2.group(3)}-{int(m2.group(1)):02d}-{int(m2.group(2)):02d}"
    return None

def find_best_header_row(file_obj, req_cols):
    """尋找標題列所在的列數"""
    try:
        df_p = pd.read_excel(file_obj, header=None, nrows=25)
    except Exception:
        file_obj.seek(0)
        df_p = pd.read_csv(file_obj, header=None, nrows=25)
    
    best_row = 0
    max_m = -1
    for i in range(min(15, len(df_p))):
        row_vals = [str(x).strip() for x in df_p.iloc[i].dropna()]
        matches = sum(1 for col in req_cols if any(col in val for val in row_vals))
        if matches > max_m:
            max_m = matches
            best_row = i
    file_obj.seek(0)
    return best_row

# 檔案上傳
uploaded_files = st.sidebar.file_uploader("📂 請上傳銷貨明細/交易檔案", type=["xls", "xlsx", "csv"], accept_multiple_files=True)

if uploaded_files:
    try:
        all_dfs = []
        req_headers = ['客戶', '商品', '金額', '數量', '單號', '日期', '載具']

        for f_obj in uploaded_files:
            f_obj.seek(0)
            skip_r = find_best_header_row(f_obj, req_headers)
            
            if f_obj.name.endswith('.csv'):
                df_curr = pd.read_csv(f_obj, skiprows=skip_r)
            else:
                df_curr = pd.read_excel(f_obj, skiprows=skip_r)
                
            df_curr.columns = df_curr.columns.astype(str).str.strip()
            all_dfs.append(df_curr)

        full_df = pd.concat(all_dfs, ignore_index=True)
        st.sidebar.success(f"已讀取 {len(uploaded_files)} 個檔案，共 {len(full_df)} 筆原始記錄。")

        # 欄位智慧映射對齊
        cols = list(full_df.columns)
        def match_col(keywords, default_val=None):
            for kw in keywords:
                for c in cols:
                    if kw.lower() in c.lower():
                        return c
            return default_val

        col_id = match_col(['客戶編號', '客戶代號', '會員編號', '卡號', '客戶'], '客戶編號')
        col_item = match_col(['商品名稱', '品名', '商品', '項目'], '商品名稱')
        col_qty = match_col(['數量', '銷貨數量', '重量', '件數'], '數量')
        col_amount = match_col(['小計', '金額', '銷貨金額', '總計', '含稅金額'], '金額')
        col_date = match_col(['銷貨日', '單據日期', '日期', '交易日'], '銷貨日')
        col_order = match_col(['銷貨單號', '單號', '交易序號', '發票號碼'], None)
        col_carrier = match_col(['載具', '共通載具', '載具號碼', '手機條碼', '發票載具'], None)

        # 數值與格式清洗
        if col_id in full_df.columns:
            full_df['客戶編號_清洗'] = full_df[col_id].apply(clean_customer_id)
        else:
            full_df['客戶編號_清洗'] = "未知客戶"

        if col_date in full_df.columns:
            full_df['標準日期'] = full_df[col_date].apply(try_parse_date)
        else:
            full_df['標準日期'] = None

        if col_amount in full_df.columns:
            full_df['標準金額'] = pd.to_numeric(full_df[col_amount].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        else:
            full_df['標準金額'] = 0.0

        if col_qty in full_df.columns:
            full_df['標準數量'] = pd.to_numeric(full_df[col_qty].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
        else:
            full_df['標準數量'] = 0.0

        # 判斷有無載具
        if col_carrier and col_carrier in full_df.columns:
            full_df['載具狀態'] = full_df[col_carrier].apply(
                lambda x: "有載具" if pd.notnull(x) and str(x).strip() not in ['', 'nan', '無', 'None', '0'] else "無載具"
            )
            full_df['載具內容'] = full_df[col_carrier].fillna('')
        else:
            full_df['載具狀態'] = "未提供欄位"
            full_df['載具內容'] = ""

        # --- 側邊欄：多模式篩選區 ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("🎯 篩選模式")

        # 1. 客戶搜尋 (關鍵字搜尋或下拉)
        all_customers = sorted([c for c in full_df['客戶編號_清洗'].unique() if c])
        search_kw = st.sidebar.text_input("🔍 客戶編號快速搜尋", placeholder="輸入如：105 或 完整代碼")

        matched_customers = [c for c in all_customers if search_kw.strip() in c] if search_kw else all_customers
        selected_customer = st.sidebar.selectbox("選擇要分析的客戶", options=["(全部客戶)"] + matched_customers)

        # 2. 日期篩選
        dates_available = sorted([d for d in full_df['標準日期'].dropna().unique()])
        if dates_available:
            date_range = st.sidebar.date_input(
                "📅 交易日期區間",
                value=(datetime.strptime(dates_available[0], '%Y-%m-%d'), datetime.strptime(dates_available[-1], '%Y-%m-%d'))
            )
        else:
            date_range = None

        # 3. 載具篩選
        carrier_options = ["全部", "有載具", "無載具"]
        carrier_filter = st.sidebar.radio("💳 載具條件", carrier_options, horizontal=True)

        # 4. 品項搜尋
        item_kw = st.sidebar.text_input("🥩 商品關鍵字篩選", placeholder="如：五花肉、梅花排")

        # --- 執行篩選 ---
        filtered_df = full_df.copy()

        if selected_customer != "(全部客戶)":
            filtered_df = filtered_df[filtered_df['客戶編號_清洗'] == selected_customer]

        if date_range and len(date_range) == 2:
            start_d, end_d = date_range[0].strftime('%Y-%m-%d'), date_range[1].strftime('%Y-%m-%d')
            filtered_df = filtered_df[(filtered_df['標準日期'] >= start_d) & (filtered_df['標準日期'] <= end_d)]

        if carrier_filter != "全部" and '載具狀態' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['載具狀態'] == carrier_filter]

        if item_kw.strip() and col_item in filtered_df.columns:
            filtered_df = filtered_df[filtered_df[col_item].astype(str).str.contains(item_kw.strip(), na=False)]

        # --- 結果展示 ---
        st.subheader(f"📊 查詢結果概覽：{selected_customer}")

        total_amount = filtered_df['標準金額'].sum()
        total_quantity = filtered_df['標準數量'].sum()
        record_count = len(filtered_df)

        order_count = filtered_df[col_order].nunique() if (col_order and col_order in filtered_df.columns) else record_count
        carrier_present = (filtered_df['載具狀態'] == "有載具").any()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("💰 篩選範圍總金額", f"${total_amount:,.0f} 元")
        c2.metric("📦 總購買數量 / 重量", f"{total_quantity:,.1f}")
        c3.metric("🧾 交易筆數 / 單數", f"{record_count} 筆明細 / {order_count} 筆單據")
        c4.metric("📱 載具使用情形", "有載具" if carrier_present else "無載具 / 未登記")

        # 品項彙整表 (該客戶買了哪些品項、總計各多少)
        if col_item in filtered_df.columns and len(filtered_df) > 0:
            st.markdown("### 📋 購買品項彙整清單")
            item_summary = filtered_df.groupby(col_item).agg(
                總數量=('標準數量', 'sum'),
                總金額=('標準金額', 'sum'),
                購買次數=('標準金額', 'count')
            ).reset_index().sort_values(by='總金額', ascending=False)

            st.dataframe(
                item_summary.style.format({'總數量': '{:,.1f}', '總金額': '${:,.0f}', '購買次數': '{:d}'}),
                use_container_width=True
            )

        # 詳細交易記錄表格
        st.markdown("### 📝 詳細明細數據")
        display_cols = []
        for candidate in [col_date, col_order, '客戶編號_清洗', col_item, '標準數量', '標準金額', '載具狀態', '載具內容']:
            if candidate and candidate in filtered_df.columns and candidate not in display_cols:
                display_cols.append(candidate)

        st.dataframe(filtered_df[display_cols], use_container_width=True)

        # 匯出結果按鈕
        csv_export = filtered_df[display_cols].to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 匯出當前篩選明細為 CSV",
            data=csv_export,
            file_name=f"客戶查詢結果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ 解析與篩選時發生錯誤：{e}")
else:
    st.info("👆 請從左側側邊欄上傳 Excel 或 CSV 檔案開始查詢。")
