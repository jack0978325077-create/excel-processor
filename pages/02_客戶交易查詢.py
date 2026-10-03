import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="客戶交易與明細深度查詢工具", layout="wide", page_icon="🔍")

st.title("🔍 客戶交易與發票收據深度查詢系統")
st.write("💡 **發票總表檢視**：一筆發票/交易呈現為一橫列。點選該筆發票，即可彈出完整購買品項清單。")

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

# 彈出視窗：顯示點選發票的內部所有品項明細
@st.dialog("🧾 發票消費明細收據")
def show_order_detail_dialog(invoice_id, full_source_df):
    items = full_source_df[full_source_df['發票單號_鍵'] == invoice_id].copy()
    if items.empty:
        st.warning("查無此單詳細品項。")
        return

    first = items.iloc[0]
    total_price = items['標準金額'].sum()
    total_qty = items['標準數量'].sum()

    st.markdown(f"### 🧾 發票/單據號碼：`{invoice_id}`")
    
    col1, col2 = st.columns(2)
    col1.write(f"📅 **消費日期**：{first.get('標準日期', '無')}")
    col1.write(f"⏰ **消費時間**：{first.get('銷貨時間', '未記載')}")
    col1.write(f"👤 **客戶編號**：{first.get('客戶編號_清洗', '未知')}")

    col2.write(f"💳 **付款方式**：**{first.get('標準付款方式', '現金')}**")
    col2.write(f"📱 **載具狀態**：{first.get('載具狀態', '無')} ({first.get('載具內容', '')})")
    col2.write(f"💰 **整單總金額**：**:red[${total_price:,.0f} 元]** (共 {total_qty:.1f} 單位)")

    st.divider()
    st.markdown("#### 🥩 該發票購買品項清單")

    item_table = items[['標準品名', '標準數量', '標準金額']].rename(
        columns={'標準品名': '商品名稱', '標準數量': '數量/重量', '標準金額': '金額小計'}
    ).reset_index(drop=True)
    
    st.dataframe(
        item_table.style.format({'數量/重量': '{:,.2f}', '金額小計': '${:,.0f}'}),
        use_container_width=True
    )
    
    if pd.notnull(first.get('B_備註')) and str(first.get('B_備註')).strip() != "":
        st.caption(f"📝 備註資訊：{first['B_備註']}")

# 側邊欄上傳區
uploaded_files = st.sidebar.file_uploader("📂 請上傳 A 檔 (明細) 與 B 檔 (銷貨單查詢)", type=["xls", "xlsx", "csv"], accept_multiple_files=True)

if uploaded_files:
    try:
        detail_dfs = [] # A 檔
        time_dfs = []   # B 檔

        for f_obj in uploaded_files:
            f_obj.seek(0)
            try:
                df_preview = pd.read_excel(f_obj, header=None, nrows=5)
            except Exception:
                f_obj.seek(0)
                df_preview = pd.read_csv(f_obj, header=None, nrows=5)
            
            preview_str = str(df_preview.values)
            
            if '商品類別' in preview_str or '商品名稱' in preview_str or '品名' in preview_str:
                req_a = ['銷貨日', '數量', '客戶編號', '商品名稱', '小計', '金額', '發票']
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
            st.error("❌ 找不到銷貨明細 (A 檔)，請確認上傳檔案。")
            st.stop()

        # 1. 整理 A 檔
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
        col_inv = match_col_a(['發票號碼', '發票', '銷貨單號', '單號', '交易序號'], None)
        col_carrier = match_col_a(['載具', '共通載具', '手機條碼'], None)

        full_detail_df['客戶編號_清洗'] = full_detail_df[col_id].apply(clean_customer_id) if col_id in full_detail_df.columns else ""
        full_detail_df['標準日期'] = full_detail_df[col_date].apply(try_parse_date) if col_date in full_detail_df.columns else None
        full_detail_df['標準品名'] = full_detail_df[col_item].astype(str).str.strip() if col_item in full_detail_df.columns else ""
        
        full_detail_df['標準金額'] = pd.to_numeric(full_detail_df[col_amount].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0) if col_amount in full_detail_df.columns else 0.0
        full_detail_df['標準數量'] = pd.to_numeric(full_detail_df[col_qty].astype(str).str.replace(',', ''), errors='coerce').fillna(0.0) if col_qty in full_detail_df.columns else 0.0

        # 發票號碼或單號判定
        if col_inv and col_inv in full_detail_df.columns and full_detail_df[col_inv].dropna().nunique() > 1:
            full_detail_df['發票單號_鍵'] = full_detail_df[col_inv].astype(str).str.strip()
        else:
            full_detail_df['發票單號_鍵'] = full_detail_df['標準日期'].astype(str) + "_" + full_detail_df['客戶編號_清洗']

        # 2. 整理 B 檔（查找依據）
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

        # 3. 橫向對照（僅將 B 檔時間與備註貼入 A 檔）
        if not full_b_df.empty:
            merged_items_df = pd.merge(full_detail_df, full_b_df[['客戶編號_清洗', 'B_時間', 'B_備註']], on='客戶編號_清洗', how='left')
            merged_items_df['銷貨時間'] = merged_items_df['B_時間'].fillna('未對齊')
            merged_items_df['B_備註'] = merged_items_df['B_備註'].fillna('')
        else:
            merged_items_df = full_detail_df.copy()
            merged_items_df['銷貨時間'] = '未對齊'
            merged_items_df['B_備註'] = ''

        merged_items_df['小時'] = merged_items_df['銷貨時間'].apply(extract_hour)

        def check_pay(text):
            s = str(text)
            if any(k in s for k in ['信用卡', '卡號', '授權碼', '刷卡']):
                return "刷卡"
            return "現金"

        merged_items_df['標準付款方式'] = merged_items_df['B_備註'].apply(check_pay)

        if col_carrier and col_carrier in merged_items_df.columns:
            merged_items_df['載具狀態'] = merged_items_df[col_carrier].apply(
                lambda x: "有載具" if pd.notnull(x) and str(x).strip() not in ['', 'nan', '無', 'None', '0'] else "無載具"
            )
            merged_items_df['載具內容'] = merged_items_df[col_carrier].fillna('')
        else:
            merged_items_df['載具狀態'] = "未提供欄位"
            merged_items_df['載具內容'] = ""

        # --- 核心關鍵：將明細聚合成「一筆發票是一橫列」的總表 ---
        invoice_summary_df = merged_items_df.groupby('發票單號_鍵').agg(
            標準日期=('標準日期', 'first'),
            銷貨時間=('銷貨時間', 'first'),
            小時=('小時', 'first'),
            客戶編號_清洗=('客戶編號_清洗', 'first'),
            發票總金額=('標準金額', 'sum'),
            購買總數量=('標準數量', 'sum'),
            品項種類數=('標準品名', 'nunique'),
            全部品項清單=('標準品名', lambda s: '、'.join([x for x in s.unique() if x])),
            標準付款方式=('標準付款方式', 'first'),
            載具狀態=('載具狀態', 'first'),
            載具內容=('載具內容', 'first')
        ).reset_index()

        st.sidebar.success(f"✅ 對齊成功！共有 {len(invoice_summary_df)} 筆發票/交易記錄。")

        # --- 側邊欄篩選 ---
        st.sidebar.markdown("---")
        st.sidebar.subheader("🎯 篩選條件")

        # 客戶快速篩選
        all_customers = sorted([c for c in invoice_summary_df['客戶編號_清洗'].unique() if c])
        search_kw = st.sidebar.text_input("🔍 客戶編號搜尋", placeholder="輸入客戶編號")
        matched_customers = [c for c in all_customers if search_kw.strip() in c] if search_kw else all_customers
        selected_customer = st.sidebar.selectbox("選擇指定客戶", options=["(全部客戶)"] + matched_customers)

        # 發票總金額門檻
        max_inv_amount = float(invoice_summary_df['發票總金額'].max()) if not invoice_summary_df.empty else 10000.0
        min_price = st.sidebar.number_input("💰 發票消費總額高於 (元) :", min_value=0.0, max_value=max_inv_amount, value=0.0, step=100.0)

        # 付款方式
        selected_pay = st.sidebar.radio("💳 付款方式", ["全部", "現金", "刷卡"], horizontal=True)

        # 時段
        time_slot_options = [
            "全部時段", "早市/清晨 (05:00 - 08:59)", "上午時段 (09:00 - 11:59)",
            "中午時段 (12:00 - 13:59)", "下午時段 (14:00 - 17:59)", "晚間時段 (18:00 - 21:59)", "自訂時段區間"
        ]
        selected_time_slot = st.sidebar.selectbox("⏰ 消費時段", options=time_slot_options)
        custom_hour_range = None
        if selected_time_slot == "自訂時段區間":
            custom_hour_range = st.sidebar.slider("選擇時段 (小時)", 0, 23, (6, 18))

        # 日期
        dates_available = sorted([d for d in invoice_summary_df['標準日期'].dropna().unique()])
        date_range = st.sidebar.date_input(
            "📅 交易日期區間",
            value=(datetime.strptime(dates_available[0], '%Y-%m-%d'), datetime.strptime(dates_available[-1], '%Y-%m-%d'))
        ) if dates_available else None

        # 載具
        carrier_filter = st.sidebar.radio("📱 載具條件", ["全部", "有載具", "無載具"], horizontal=True)

        # 品項關鍵字
        st.sidebar.markdown("---")
        item_kw = st.sidebar.text_input("🥩 品項關鍵字搜尋 (例如：軟骨)", value="", placeholder="輸入軟骨、梅花肉...")

        # --- 執行發票總表篩選 ---
        filtered_inv_df = invoice_summary_df.copy()

        if selected_customer != "(全部客戶)":
            filtered_inv_df = filtered_inv_df[filtered_inv_df['客戶編號_清洗'] == selected_customer]

        if min_price > 0:
            filtered_inv_df = filtered_inv_df[filtered_inv_df['發票總金額'] >= min_price]

        if selected_pay != "全部":
            filtered_inv_df = filtered_inv_df[filtered_inv_df['標準付款方式'] == selected_pay]

        if selected_time_slot != "全部時段" and '小時' in filtered_inv_df.columns:
            slots = {
                "早市/清晨 (05:00 - 08:59)": (5, 8),
                "上午時段 (09:00 - 11:59)": (9, 11),
                "中午時段 (12:00 - 13:59)": (12, 13),
                "下午時段 (14:00 - 17:59)": (14, 17),
                "晚間時段 (18:00 - 21:59)": (18, 21),
            }
            if selected_time_slot in slots:
                r = slots[selected_time_slot]
                filtered_inv_df = filtered_inv_df[filtered_inv_df['小時'].between(r[0], r[1])]
            elif selected_time_slot == "自訂時段區間" and custom_hour_range:
                filtered_inv_df = filtered_inv_df[filtered_inv_df['小時'].between(custom_hour_range[0], custom_hour_range[1])]

        if date_range and len(date_range) == 2:
            start_d, end_d = date_range[0].strftime('%Y-%m-%d'), date_range[1].strftime('%Y-%m-%d')
            filtered_inv_df = filtered_inv_df[(filtered_inv_df['標準日期'] >= start_d) & (filtered_inv_df['標準日期'] <= end_d)]

        if carrier_filter != "全部":
            filtered_inv_df = filtered_inv_df[filtered_inv_df['載具狀態'] == carrier_filter]

        # 關鍵字過濾：只要該發票的品項清單中有包含關鍵字就留下
        clean_kw_list = [k.strip() for k in re.split(r'[,，\s]+', item_kw) if k.strip()]
        if clean_kw_list:
            kw_pattern = '|'.join([re.escape(k) for k in clean_kw_list])
            filtered_inv_df = filtered_inv_df[filtered_inv_df['全部品項清單'].str.contains(kw_pattern, case=False, na=False)]

        # --- 數據呈現 ---
        st.subheader("📊 發票查詢結果概況")

        if len(filtered_inv_df) == 0:
            st.warning("⚠️ 目前查無任何發票！請檢查付款方式（是否誤選刷卡但為現金）或將金額門檻調低。")
        else:
            total_amount = filtered_inv_df['發票總金額'].sum()
            total_quantity = filtered_inv_df['購買總數量'].sum()
            total_invoices = len(filtered_inv_df)
            has_carrier_count = (filtered_inv_df['載具狀態'] == "有載具").sum()
            carrier_rate = (has_carrier_count / total_invoices * 100) if total_invoices > 0 else 0
            cash_count = (filtered_inv_df['標準付款方式'] == "現金").sum()
            card_count = (filtered_inv_df['標準付款方式'] == "刷卡").sum()

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("💰 符合發票總金額", f"${total_amount:,.0f} 元")
            c2.metric("🧾 符合發票數", f"{total_invoices} 張")
            c3.metric("💳 現金 / 刷卡單數", f"現金 {cash_count} 張 | 刷卡 {card_count} 張")
            c4.metric("📱 載具發票比例", f"{carrier_rate:.1f}% ({has_carrier_count} 張)")

            # 發票總表（一筆發票是一橫列）
            st.markdown("### 📋 發票清單（💡 點擊任一發票列最左側的圓圈，即可彈出該發票完整明細收據）")

            display_cols = ['標準日期', '銷貨時間', '發票單號_鍵', '客戶編號_清洗', '發票總金額', '購買總數量', '標準付款方式', '載具狀態', '全部品項清單']
            table_show = filtered_inv_df[display_cols].reset_index(drop=True)

            event = st.dataframe(
                table_show.style.format({'發票總金額': '${:,.0f}', '購買總數量': '{:,.2f}'}),
                use_container_width=True,
                on_select="rerun",
                selection_mode="single-row"
            )

            # 點選彈窗
            selected_rows = event.selection.rows if hasattr(event, 'selection') else []
            if len(selected_rows) > 0:
                selected_idx = selected_rows[0]
                clicked_inv_id = table_show.iloc[selected_idx]['發票單號_鍵']
                show_order_detail_dialog(clicked_inv_id, merged_items_df)

            # 匯出 CSV 按鈕
            csv_export = table_show.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 匯出當前發票總表為 CSV",
                data=csv_export,
                file_name=f"發票總表查詢_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"❌ 查詢處理時發生錯誤：{e}")
else:
    st.info("👆 請從左側側邊欄同時上傳 A 檔 (明細) 與 B 檔 (銷貨單查詢)。")
