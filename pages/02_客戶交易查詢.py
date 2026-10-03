import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="客戶交易與明細深度查詢工具", layout="wide", page_icon="🔍")

st.title("🔍 客戶交易與品項明細深度查詢系統")
st.write("💡 **A 檔 (明細檔)** 提供品項與金額，**B 檔 (銷貨單查詢)** 僅做時間與刷卡備註對照。支援點擊任一列彈出完整交易收據。")

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

# 彈出視窗：顯示該筆單據完整明細
@st.dialog("🧾 交易詳細明細收據")
def show_order_detail_dialog(order_key, full_merged_df):
    order_items = full_merged_df[full_merged_df['交易對齊鍵'] == order_key].copy()
    if order_items.empty:
        st.warning("查無該單詳細內容。")
        return

    first_row = order_items.iloc[0]
    total_order_price = order_items['標準金額'].sum()
    total_order_qty = order_items['標準數量'].sum()

    st.markdown(f"### 📍 客戶編號：`{first_row['客戶編號_清洗']}`")
    
    m1, m2 = st.columns(2)
    m1.write(f"📅 **消費日期**：{first_row.get('標準日期', '無')}")
    m1.write(f"⏰ **消費時間**：{first_row.get('銷貨時間', '未記載')}")
    
    m2.write(f"💳 **付款方式**：**{first_row.get('標準付款方式', '現金')}**")
    m2.write(f"📱 **載具狀態**：{first_row.get('載具狀態', '無')} ({first_row.get('載具內容', '')})")
    m2.write(f"💰 **整單總額 (A檔品項加總)**：**:red[${total_order_price:,.0f} 元]** (共 {total_order_qty:.1f} 單位)")

    st.divider()
    st.markdown("#### 🥩 本筆交易包含的所有商品品項 (來自 A 檔)")

    receipt_table = order_items[['標準品名', '標準數量', '標準金額']].rename(
        columns={'標準品名': '商品名稱', '標準數量': '數量/重量', '標準金額': '金額小計'}
    ).reset_index(drop=True)
    
    st.dataframe(
        receipt_table.style.format({'數量/重量': '{:,.1f}', '金額小計': '${:,.0f}'}),
        use_container_width=True
    )
    
    if pd.notnull(first_row.get('B_備註')) and str(first_row.get('B_備註')).strip() != "":
        st.caption(f"📝 B 檔備註對照：{first_row['B_備註']}")

# 側邊欄上傳檔案
uploaded_files = st.sidebar.file_uploader("📂 請同時上傳 A 檔 (明細) 與 B 檔 (銷貨單查詢)", type=["xls", "xlsx", "csv"], accept_multiple_files=True)

if uploaded_files:
    try:
        detail_dfs = [] # A 檔 (提供品項、金額、數量)
        time_dfs = []   # B 檔 (僅提供時間與刷卡備註對照)

        for f_obj in uploaded_files:
            f_obj.seek(0)
            try:
                df_preview = pd.read_excel(f_obj, header=None, nrows=5)
            except Exception:
                f_obj.seek(0)
                df_preview = pd.read_csv(f_obj, header=None, nrows=5)
            
            preview_str = str(df_preview.values)
            
            # 依欄位特徵分流 A 檔與 B 檔
            if '商品類別' in preview_str or '商品名稱' in preview_str or '品名' in preview_str:
                req_a = ['銷貨日', '數量', '客戶編號', '商品名稱', '小計', '金額']
                f_obj.seek(0)
                try:
                    df_p = pd.read_excel(f_obj, header=None, nrows=20)
                except Exception:
                    df_p = pd.read_csv(f_obj, header=None, nrows=20)
                
                best_row = 0
                max_m = -1
                for i in range(min(10, len(df_p))):
                    row_v = [str(x).strip() for x in df_p.iloc[i].dropna()]
                    m = sum(1 for col in req_a if any(col in val for val in row_v))
                    if m > max_m:
                        max_m = m
                        best_row = i
                
                f_obj.seek(0)
                df_actual = pd.read_csv(f_obj, skiprows=best_row) if f_obj.name.endswith('.csv') else pd.read_excel(f_obj, skiprows=best_row)
                df_actual.columns = df_actual.columns.astype(str).str.strip()
                detail_dfs.append(df_actual)
            else:
                f_obj.seek(0)
                df_b_raw = pd.read_csv(f_obj, header=None) if f_obj.name.endswith('.csv') else pd.read_excel(f_obj, header=None)
                time_dfs.append(df_b_raw)

        if len(detail_dfs) == 0:
            st.error("❌ 找不到任何銷貨明細 (A 檔)，請確認上傳的檔案。")
            st.stop()

        # 1. 解析 A 檔（所有金額與品項資料的唯一來源）
        full_detail_df = pd.concat(detail_dfs, ignore_index=True)
        cols_a = list(full_detail_df.columns)

        def match_col_a(keywords, default_val=None):
            for kw in keywords:
                for c in cols_a:
                    if kw.lower() in c.lower():
                        return c
            return default_val

        col_id = match_col_a(['客戶編號', '客戶代號', '會員編號', '客戶'], '客戶編號')
        col_item = match_col_a(['商品名稱', '品名', '商品'], '商品名稱')
        col_qty = match_col_a(['數量', '銷貨數量', '重量'], '數量')
        col_amount = match_col_a(['小計', '金額', '銷貨金額', '總計'], '金額')
        col_date = match_col_a(['銷貨日', '日期', '交易日'], '銷貨日')
        col_carrier = match_col_a(['載具', '共通載具', '手機條碼'], None)

        full_detail_df['客戶編號_清洗'] = full_detail_df[col_id].apply(clean_customer_id) if col_id in full_detail_df.columns else ""
        full_detail_df['標準日期'] = full_detail_df[col_date].apply(try_parse_date) if col_date in full_detail_df.columns else None
        full_detail_df['標準品名'] = full_detail_df[col_item].astype(str).str.strip() if col_item in full_detail_df.columns else ""
        
        # 單一商品項目的金額與數量（僅取自 A 檔）
        full_detail_df['標準金額'] = pd.to_numeric(full_detail_df[col_amount].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0) if col_amount in full_detail_df.columns else 0.0
        full_detail_df['標準數量'] = pd.to_numeric(full_detail_df[col_qty].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0) if col_qty in full_detail_df.columns else 0.0

        # 以「日期 + 客戶」定義一筆交易
        full_detail_df['交易對齊鍵'] = full_detail_df['標準日期'].astype(str) + "_" + full_detail_df['客戶編號_清洗']

        # 計算 A 檔該客戶整單總額（純粹 A 檔各項小計加總）
        order_total_map = full_detail_df.groupby('交易對齊鍵')['標準金額'].sum().to_dict()
        full_detail_df['整單總金額'] = full_detail_df['交易對齊鍵'].map(order_total_map)

        # 2. 解析 B 檔（僅做時間與備註查找依據，絕不抓取金額）
        combined_b_list = []
        for df_b in time_dfs:
            num_cols = len(df_b.columns)
            if num_cols >= 8:
                col_h = df_b[7].astype(str).str.strip()
                col_i = df_b[8].astype(str).str.strip() if num_cols >= 9 else pd.Series()
                
                if not col_i.empty and col_i.str.contains(':', na=False).sum() > col_h.str.contains(':', na=False).sum():
                    cust_idx, time_idx = 7, 8
                else:
                    cust_idx, time_idx = 8 if num_cols >= 9 else 0, 7

                remark_idx = None
                for c_idx in range(num_cols):
                    col_str = df_b[c_idx].astype(str)
                    if col_str.str.contains('信用卡|卡號|授權碼', na=False).sum() > 0:
                        remark_idx = c_idx
                        break

                b_sub = pd.DataFrame()
                b_sub['B_客戶'] = df_b[cust_idx]
                b_sub['B_時間'] = df_b[time_idx]
                b_sub['B_備註'] = df_b[remark_idx] if remark_idx is not None else ""
                b_sub['客戶編號_清洗'] = b_sub['B_客戶'].apply(clean_customer_id)
                combined_b_list.append(b_sub)

        if combined_b_list:
            full_b_df = pd.concat(combined_b_list, ignore_index=True)
            full_b_df = full_b_df.dropna(subset=['客戶編號_清洗']).drop_duplicates(subset=['客戶編號_清洗'])
        else:
            full_b_df = pd.DataFrame(columns=['客戶編號_清洗', 'B_時間', 'B_備註'])

        # 3. 橫向對照（僅把 B 檔的時間與備註貼到 A 檔）
        if not full_b_df.empty:
            merged_df = pd.merge(full_detail_df, full_b_df[['客戶編號_清洗', 'B_時間', 'B_備註']], on='客戶編號_清洗', how='left')
            merged_df['銷貨時間'] = merged_df['B_時間'].fillna('未對齊')
            merged_df['B_備註'] = merged_df['B_備註'].fillna('')
        else:
            merged_df = full_detail_df.copy()
            merged_df['銷貨時間'] = '未對齊'
            merged_df['B_備註'] = ''

        merged_df['小時'] = merged_df['銷貨時間'].apply(extract_hour)

        # 判斷付款方式（純看 B 檔備註是否含有信用卡特徵）
        def check_pay(text):
            s = str(text)
            if any(k in s for k in ['信用卡', '卡號', '授權碼', '刷卡']):
                return "刷卡"
            return "現金"

        merged_df['標準付款方式'] = merged_df['B_備註'].apply(check_pay)

        # 載具判斷
        if col_carrier and col_carrier in merged_df.columns:
            merged_df['載具狀態'] = merged_df[col_carrier].apply(
                lambda x: "有載具" if pd.notnull(x) and str(x).strip() not in ['', 'nan', '無', 'None', '0'] else "無載具"
            )
            merged_df['載具內容'] = merged_df[col_carrier].fillna('')
        else:
            merged_df['載具狀態'] = "未提供欄位"
            merged_df['載具內容'] = ""

        st.sidebar.success(f"✅ 對齊完成！共 {len(merged_df)} 筆明細，金額純取自 A 檔。")

        # --- 側邊欄篩選 ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("🎯 篩選條件")

        # 客戶編號快速搜尋
        all_customers = sorted([c for c in merged_df['客戶編號_清洗'].unique() if c])
        search_kw = st.sidebar.text_input("🔍 客戶編號搜尋", placeholder="輸入代碼或編號")
        matched_customers = [c for c in all_customers if search_kw.strip() in c] if search_kw else all_customers
        selected_customer = st.sidebar.selectbox("選擇指定客戶", options=["(全部客戶)"] + matched_customers)

        # 金額門檻篩選（基準純為 A 檔商品加總）
        price_mode = st.sidebar.radio("金額門檻計算方式：", ["整筆交易總額 (A檔同單加總)", "單一品項金額"], index=0)
        max_amount_data = float(merged_df['整單總金額'].max()) if not merged_df.empty else 10000.0
        min_price = st.sidebar.number_input("💰 消費金額高於 (元) :", min_value=0.0, max_value=max_amount_data, value=0.0, step=100.0)

        # 付款方式（依 B 檔備註識別）
        selected_pay = st.sidebar.radio("💳 付款方式", ["全部", "現金", "刷卡"], horizontal=True)

        # 消費時段
        time_slot_options = [
            "全部時段", "早市/清晨 (05:00 - 08:59)", "上午時段 (09:00 - 11:59)",
            "中午時段 (12:00 - 13:59)", "下午時段 (14:00 - 17:59)", "晚間時段 (18:00 - 21:59)", "自訂時段區間"
        ]
        selected_time_slot = st.sidebar.selectbox("⏰ 消費時段", options=time_slot_options)
        custom_hour_range = None
        if selected_time_slot == "自訂時段區間":
            custom_hour_range = st.sidebar.slider("選擇時段 (小時)", 0, 23, (6, 18))

        # 日期區間
        dates_available = sorted([d for d in merged_df['標準日期'].dropna().unique()])
        date_range = st.sidebar.date_input(
            "📅 交易日期區間",
            value=(datetime.strptime(dates_available[0], '%Y-%m-%d'), datetime.strptime(dates_available[-1], '%Y-%m-%d'))
        ) if dates_available else None

        # 載具條件
        carrier_filter = st.sidebar.radio("📱 載具條件", ["全部", "有載具", "無載具"], horizontal=True)

        # 品項關鍵字
        st.sidebar.markdown("---")
        st.sidebar.subheader("🥩 品項關鍵字設定")
        item_kw = st.sidebar.text_input("輸入關鍵字 (如：軟骨 或 軟骨,五花)", value="", placeholder="例如：軟骨")
        kw_mode = st.sidebar.radio(
            "關鍵字呈現範圍：",
            ["反查整筆交易（有買此商品的那整單都要看）", "僅看含有關鍵字的商品明細"],
            index=0
        )

        # --- 資料過濾 ---
        filtered_df = merged_df.copy()

        if selected_customer != "(全部客戶)":
            filtered_df = filtered_df[filtered_df['客戶編號_清洗'] == selected_customer]

        if min_price > 0:
            if "整筆交易總額" in price_mode:
                filtered_df = filtered_df[filtered_df['整單總金額'] >= min_price]
            else:
                filtered_df = filtered_df[filtered_df['標準金額'] >= min_price]

        if selected_pay != "全部":
            filtered_df = filtered_df[filtered_df['標準付款方式'] == selected_pay]

        if selected_time_slot != "全部時段" and '小時' in filtered_df.columns:
            slots = {
                "早市/清晨 (05:00 - 08:59)": (5, 8),
                "上午時段 (09:00 - 11:59)": (9, 11),
                "中午時段 (12:00 - 13:59)": (12, 13),
                "下午時段 (14:00 - 17:59)": (14, 17),
                "晚間時段 (18:00 - 21:59)": (18, 21),
            }
            if selected_time_slot in slots:
                r = slots[selected_time_slot]
                filtered_df = filtered_df[filtered_df['小時'].between(r[0], r[1])]
            elif selected_time_slot == "自訂時段區間" and custom_hour_range:
                filtered_df = filtered_df[filtered_df['小時'].between(custom_hour_range[0], custom_hour_range[1])]

        if date_range and len(date_range) == 2:
            start_d, end_d = date_range[0].strftime('%Y-%m-%d'), date_range[1].strftime('%Y-%m-%d')
            filtered_df = filtered_df[(filtered_df['標準日期'] >= start_d) & (filtered_df['標準日期'] <= end_d)]

        if carrier_filter != "全部" and '載具狀態' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['載具狀態'] == carrier_filter]

        clean_kw_list = [k.strip() for k in re.split(r'[,，\s]+', item_kw) if k.strip()]
        if clean_kw_list:
            kw_pattern = '|'.join([re.escape(k) for k in clean_kw_list])
            has_kw_mask = filtered_df['標準品名'].str.contains(kw_pattern, case=False, na=False)

            if kw_mode == "僅看含有關鍵字的商品明細":
                filtered_df = filtered_df[has_kw_mask]
            else:
                matched_orders = filtered_df[has_kw_mask]['交易對齊鍵'].unique()
                filtered_df = filtered_df[filtered_df['交易對齊鍵'].isin(matched_orders)]

        # --- 數據呈現 ---
        st.subheader("📊 篩選查詢成果概況")

        if len(filtered_df) == 0:
            st.warning("⚠️ 目前查無資料！請確認篩選條件是否過於嚴格。")
        else:
            total_amount = filtered_df['標準金額'].sum()
            total_quantity = filtered_df['標準數量'].sum()
            record_count = len(filtered_df)
            order_count = filtered_df['交易對齊鍵'].nunique()
            has_carrier_count = (filtered_df['載具狀態'] == "有載具").sum()
            carrier_rate = (has_carrier_count / record_count * 100) if record_count > 0 else 0
            cash_count = (filtered_df['標準付款方式'] == "現金").sum()
            card_count = (filtered_df['標準付款方式'] == "刷卡").sum()

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("💰 篩選範圍總額", f"${total_amount:,.0f} 元")
            c2.metric("📦 總採購數量/重量", f"{total_quantity:,.1f}")
            c3.metric("💳 現金 / 刷卡筆數", f"現金 {cash_count} 筆 | 刷卡 {card_count} 筆")
            c4.metric("📱 載具比例", f"{carrier_rate:.1f}% ({has_carrier_count} 筆)")

            # 品項彙整表
            st.markdown("### 📋 購買品項彙整清單")
            item_summary = filtered_df.groupby('標準品名').agg(
                總數量=('標準數量', 'sum'),
                總金額=('標準金額', 'sum'),
                筆數=('標準金額', 'count')
            ).reset_index().sort_values(by='總金額', ascending=False)

            st.dataframe(
                item_summary.style.format({'總數量': '{:,.1f}', '總金額': '${:,.0f}', '筆數': '{:d}'}),
                use_container_width=True
            )

            # 詳細清單 (點擊跳收據)
            st.markdown("### 📝 詳細明細清單（💡 點擊左側圓圈勾選任一列，可彈出該筆交易完整收據）")
            
            display_cols = ['標準日期', '銷貨時間', '客戶編號_清洗', '標準品名', '標準數量', '標準金額', '整單總金額', '標準付款方式', '載具狀態', '交易對齊鍵']
            table_for_select = filtered_df[[c for c in display_cols if c in filtered_df.columns]].reset_index(drop=True)

            event = st.dataframe(
                table_for_select,
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row"
            )

            selected_rows = event.selection.rows if hasattr(event, 'selection') else []
            if len(selected_rows) > 0:
                selected_idx = selected_rows[0]
                clicked_key = table_for_select.iloc[selected_idx]['交易對齊鍵']
                show_order_detail_dialog(clicked_key, merged_df)

            # 匯出 CSV
            csv_export = table_for_select.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 匯出當前篩選結果為 CSV",
                data=csv_export,
                file_name=f"交易篩選查詢_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"❌ 查詢處理時發生錯誤：{e}")
else:
    st.info("👆 請從左側側邊欄同時上傳您的 A 檔 (明細) 與 B 檔 (銷貨單查詢)。")
