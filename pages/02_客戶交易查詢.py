import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="客戶交易與明細深度查詢工具", layout="wide", page_icon="🔍")

st.title("🔍 客戶交易與品項明細深度查詢系統")
st.write("💡 支援多條件組合查詢：**特定客戶、消費金額門檻、現金/刷卡支付方式、消費時段、載具狀態、商品關鍵字**。")

def clean_customer_id(x):
    if pd.isnull(x):
        return ""
    s = str(x).strip()
    if s.endswith('.0'):
        s = s[:-2]
    if s.isdigit():
        return str(int(s))
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

def extract_hour(time_val):
    if pd.isnull(time_val):
        return None
    s = str(time_val).strip()
    m = re.search(r'(\d{1,2}):(\d{2})', s)
    if m:
        h = int(m.group(1))
        if 0 <= h <= 23:
            return h
    return None

def find_best_header_row(file_obj, req_cols):
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

# 側邊欄上傳檔案
uploaded_files = st.sidebar.file_uploader("📂 請上傳銷貨明細 / 交易檔案 (支援多檔)", type=["xls", "xlsx", "csv"], accept_multiple_files=True)

if uploaded_files:
    try:
        all_dfs = []
        req_headers = ['客戶', '商品', '金額', '數量', '單號', '日期', '載具', '時間', '備註']

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
        st.sidebar.success(f"已讀取 {len(uploaded_files)} 個檔案，共 {len(full_df)} 筆資料。")

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
        col_time = match_col(['時間', '銷貨時間', '交易時間', '結帳時間'], None)
        col_order = match_col(['銷貨單號', '單號', '交易序號', '發票號碼'], None)
        col_carrier = match_col(['載具', '共通載具', '載具號碼', '手機條碼', '發票載具'], None)
        col_remark = match_col(['備註', '列印備註', '說明'], None)

        # 基礎欄位清洗
        full_df['客戶編號_清洗'] = full_df[col_id].apply(clean_customer_id) if col_id in full_df.columns else "未知客戶"
        full_df['標準日期'] = full_df[col_date].apply(try_parse_date) if col_date in full_df.columns else None
        
        # 標記商品名稱（轉純文字避免空值比對錯誤）
        if col_item in full_df.columns:
            full_df['標準品名'] = full_df[col_item].astype(str).str.strip()
        else:
            full_df['標準品名'] = ""

        # 單號欄位處理
        if col_order and col_order in full_df.columns:
            full_df['標準單號'] = full_df[col_order].astype(str).str.strip()
        else:
            # 若無單號欄位，以 日期+客戶 視為同筆單據
            full_df['標準單號'] = full_df['標準日期'].astype(str) + "_" + full_df['客戶編號_清洗']

        if col_time and col_time in full_df.columns:
            full_df['小時'] = full_df[col_time].apply(extract_hour)
        else:
            full_df['小時'] = None

        if col_amount in full_df.columns:
            full_df['標準金額'] = pd.to_numeric(full_df[col_amount].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
        else:
            full_df['標準金額'] = 0.0

        if col_qty in full_df.columns:
            full_df['標準數量'] = pd.to_numeric(full_df[col_qty].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0)
        else:
            full_df['標準數量'] = 0.0

        # 付款方式精準判斷
        def detect_card_or_cash(row):
            remark_text = ""
            if col_remark and col_remark in row and pd.notnull(row[col_remark]):
                remark_text += str(row[col_remark]).strip()
            if any(k in remark_text for k in ['信用卡', '卡號', '授權碼', '刷卡']):
                return "刷卡"
            return "現金"

        full_df['標準付款方式'] = full_df.apply(detect_card_or_cash, axis=1)

        # 載具狀態
        if col_carrier and col_carrier in full_df.columns:
            full_df['載具狀態'] = full_df[col_carrier].apply(
                lambda x: "有載具" if pd.notnull(x) and str(x).strip() not in ['', 'nan', '無', 'None', '0'] else "無載具"
            )
            full_df['載具內容'] = full_df[col_carrier].fillna('')
        else:
            full_df['載具狀態'] = "未提供欄位"
            full_df['載具內容'] = ""

        # --- 側邊欄篩選設定 ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("🎯 篩選條件")

        # 1. 客戶編號快速搜尋
        all_customers = sorted([c for c in full_df['客戶編號_清洗'].unique() if c])
        search_kw = st.sidebar.text_input("🔍 客戶編號搜尋", placeholder="輸入代碼或編號")
        matched_customers = [c for c in all_customers if search_kw.strip() in c] if search_kw else all_customers
        selected_customer = st.sidebar.selectbox("選擇指定客戶", options=["(全部客戶)"] + matched_customers)

        # 2. 金額門檻篩選
        max_amount_data = float(full_df['標準金額'].max()) if not full_df.empty else 10000.0
        min_price = st.sidebar.number_input("💰 消費金額高於 (元) :", min_value=0.0, max_value=max_amount_data, value=0.0, step=100.0)

        # 3. 付款方式篩選
        selected_pay = st.sidebar.radio("💳 付款方式", ["全部", "現金", "刷卡"], horizontal=True)

        # 4. 消費時段篩選
        time_slot_options = [
            "全部時段",
            "早市/清晨 (05:00 - 08:59)",
            "上午時段 (09:00 - 11:59)",
            "中午時段 (12:00 - 13:59)",
            "下午時段 (14:00 - 17:59)",
            "晚間時段 (18:00 - 21:59)",
            "自訂時段區間"
        ]
        selected_time_slot = st.sidebar.selectbox("⏰ 消費時段", options=time_slot_options)

        custom_hour_range = None
        if selected_time_slot == "自訂時段區間":
            custom_hour_range = st.sidebar.slider("選擇時段 (小時)", 0, 23, (6, 18))

        # 5. 日期區間
        dates_available = sorted([d for d in full_df['標準日期'].dropna().unique()])
        if dates_available:
            date_range = st.sidebar.date_input(
                "📅 交易日期區間",
                value=(datetime.strptime(dates_available[0], '%Y-%m-%d'), datetime.strptime(dates_available[-1], '%Y-%m-%d'))
            )
        else:
            date_range = None

        # 6. 載具條件
        carrier_filter = st.sidebar.radio("📱 載具條件", ["全部", "有載具", "無載具"], horizontal=True)

        # 7. 品項關鍵字篩選 (支援多詞與整筆交易反查)
        st.sidebar.markdown("---")
        st.sidebar.subheader("🥩 品項關鍵字設定")
        item_kw = st.sidebar.text_input("輸入關鍵字 (如：軟骨 或 軟骨,五花)", placeholder="輸入如：軟骨")
        kw_mode = st.sidebar.radio(
            "關鍵字呈現範圍：",
            ["僅看含有關鍵字的商品明細", "反查整筆交易（有買此商品的那整單都要看）"],
            index=0
        )

        # --- 資料過濾 ---
        filtered_df = full_df.copy()

        if selected_customer != "(全部客戶)":
            filtered_df = filtered_df[filtered_df['客戶編號_清洗'] == selected_customer]

        if min_price > 0:
            filtered_df = filtered_df[filtered_df['標準金額'] >= min_price]

        if selected_pay != "全部":
            filtered_df = filtered_df[filtered_df['標準付款方式'] == selected_pay]

        if selected_time_slot != "全部時段" and '小時' in filtered_df.columns:
            if selected_time_slot == "早市/清晨 (05:00 - 08:59)":
                filtered_df = filtered_df[filtered_df['小時'].between(5, 8)]
            elif selected_time_slot == "上午時段 (09:00 - 11:59)":
                filtered_df = filtered_df[filtered_df['小時'].between(9, 11)]
            elif selected_time_slot == "中午時段 (12:00 - 13:59)":
                filtered_df = filtered_df[filtered_df['小時'].between(12, 13)]
            elif selected_time_slot == "下午時段 (14:00 - 17:59)":
                filtered_df = filtered_df[filtered_df['小時'].between(14, 17)]
            elif selected_time_slot == "晚間時段 (18:00 - 21:59)":
                filtered_df = filtered_df[filtered_df['小時'].between(18, 21)]
            elif selected_time_slot == "自訂時段區間" and custom_hour_range:
                filtered_df = filtered_df[filtered_df['小時'].between(custom_hour_range[0], custom_hour_range[1])]

        if date_range and len(date_range) == 2:
            start_d, end_d = date_range[0].strftime('%Y-%m-%d'), date_range[1].strftime('%Y-%m-%d')
            filtered_df = filtered_df[(filtered_df['標準日期'] >= start_d) & (filtered_df['標準日期'] <= end_d)]

        if carrier_filter != "全部" and '載具狀態' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['載具狀態'] == carrier_filter]

        # 關鍵字智能處理（切分逗點或空白）
        clean_kw_list = [k.strip() for k in re.split(r'[,，\s]+', item_kw) if k.strip()]
        if clean_kw_list:
            # 建立包含任何一個關鍵字的遮罩
            kw_pattern = '|'.join([re.escape(k) for k in clean_kw_list])
            has_kw_mask = filtered_df['標準品名'].str.contains(kw_pattern, case=False, na=False)

            if kw_mode == "僅看含有關鍵字的商品明細":
                filtered_df = filtered_df[has_kw_mask]
            else:
                # 找出符合關鍵字的所有單據編號，再把這些單據的全部商品撈出來
                matched_orders = filtered_df[has_kw_mask]['標準單號'].unique()
                filtered_df = filtered_df[filtered_df['標準單號'].isin(matched_orders)]

        # --- 統計呈現 ---
        st.subheader("📊 篩選查詢成果概況")

        total_amount = filtered_df['標準金額'].sum()
        total_quantity = filtered_df['標準數量'].sum()
        record_count = len(filtered_df)
        order_count = filtered_df['標準單號'].nunique()
        
        has_carrier_count = (filtered_df['載具狀態'] == "有載具").sum()
        carrier_rate = (has_carrier_count / record_count * 100) if record_count > 0 else 0

        cash_count = (filtered_df['標準付款方式'] == "現金").sum()
        card_count = (filtered_df['標準付款方式'] == "刷卡").sum()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("💰 總消費金額", f"${total_amount:,.0f} 元")
        c2.metric("📦 總採購數量/重量", f"{total_quantity:,.1f}")
        c3.metric("💳 現金 / 刷卡筆數", f"現金 {cash_count} 筆 | 刷卡 {card_count} 筆")
        c4.metric("📱 載具比例", f"{carrier_rate:.1f}% ({has_carrier_count} 筆有載具)")

        if len(filtered_df) > 0:
            st.markdown("### 📋 購買品項彙整清單（品名、數量、總金額）")
            item_summary = filtered_df.groupby('標準品名').agg(
                總數量=('標準數量', 'sum'),
                總金額=('標準金額', 'sum'),
                筆數=('標準金額', 'count')
            ).reset_index().sort_values(by='總金額', ascending=False)

            st.dataframe(
                item_summary.style.format({'總數量': '{:,.1f}', '總金額': '${:,.0f}', '筆數': '{:d}'}),
                use_container_width=True
            )

        # 詳細清單
        st.markdown("### 📝 詳細明細數據清單")
        display_cols = []
        candidate_cols = [col_date, col_time, col_order, '客戶編號_清洗', '標準品名', '標準數量', '標準金額', '標準付款方式', '載具狀態', '載具內容', col_remark]
        for c in candidate_cols:
            if c and c in filtered_df.columns and c not in display_cols:
                display_cols.append(c)

        st.dataframe(filtered_df[display_cols], use_container_width=True)

        # 下載按鈕
        csv_export = filtered_df[display_cols].to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 匯出當前篩選結果為 CSV",
            data=csv_export,
            file_name=f"交易篩選查詢_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"❌ 查詢處理時發生錯誤：{e}")
else:
    st.info("👆 請從左側側邊欄上傳您的銷貨檔案（支援多檔批次上傳）。")
