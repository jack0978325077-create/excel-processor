import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="超級多檔案跨天自動拆分工具", layout="centered", page_icon="🐷")

st.title("🐷 超級多檔案跨天自動拆分工具")
st.write("💡 **神級功能**：支援單一檔案內包含**多天日期混雜**。上傳後大腦會自動打散、依據內部真實日期重新分組、進行豬肉篩選與 B 檔時間校正，最後自動輸出各天的獨立 Excel。")

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
    """強大相容性的日期解析函數，能對付各種 Excel 奇怪格式，統一回傳 YYYY-MM-DD"""
    if pd.isnull(date_val):
        return None
    
    if isinstance(date_val, (datetime, pd.Timestamp)):
        return date_val.strftime('%Y-%m-%d')
        
    s = str(date_val).strip().split(' ')[0] # 去除時間部分
    
    # 格式 1: 2026-06-17 或 2026/06/17
    m1 = re.match(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
    if m1:
        return f"{m1.group(1)}-{int(m1.group(2)):02d}-{int(m1.group(3)):02d}"
        
    # 格式 2: 6/17/2026 (美式格式)
    m2 = re.match(r'(\d{1,2})[-/](\d{1,2})[-/](\d{4})', s)
    if m2:
        return f"{m2.group(3)}-{int(m2.group(1)):02d}-{int(m2.group(2)):02d}"
        
    return None

# 讓使用者一次上傳多個檔案
st.subheader("📁 檔案盲測上傳區 (支援多天份、多檔案直接拖曳)")
uploaded_files = st.file_uploader("請全選並上傳您所有的 Excel 檔案 (免改檔名、免分天)", type=["xls", "xlsx", "csv"], accept_multiple_files=True)

if uploaded_files:
    try:
        st.info(f"🔍 偵測到 {len(uploaded_files)} 個檔案，智慧大腦開始解析內部欄位與特徵...")
        
        detail_dfs = [] # 存放所有讀取出來的明細數據 (A檔數據庫)
        time_dfs = []   # 存放所有讀取出來的時間數據 (B檔數據庫)
        
        # 第一階段：辨識檔案類型並讀取內容
        for f_obj in uploaded_files:
            f_obj.seek(0)
            try:
                df_preview = pd.read_excel(f_obj, header=None, nrows=5)
            except Exception:
                f_obj.seek(0)
                df_preview = pd.read_csv(f_obj, header=None, nrows=5)
            
            preview_str = str(df_preview.values)
            
            # 以欄位特徵分流
            if '商品類別' in preview_str or '商品名稱' in preview_str:
                # 智慧尋找 A 檔標題列
                req_a = ['銷貨日', '數量', '客戶編號', '商品名稱', '商品類別']
                f_obj.seek(0)
                try:
                    df_p = pd.read_excel(f_obj, header=None, nrows=20)
                except Exception:
                    df_p = pd.read_csv(f_obj, header=None, nrows=20)
                
                best_row = 0
                max_m = -1
                for i in range(min(10, len(df_p))):
                    row_v = [str(x).strip() for x in df_p.iloc[i].dropna()]
                    m = sum(1 for col in req_a if col in row_v)
                    if m > max_m:
                        max_m = m
                        best_row = i
                
                f_obj.seek(0)
                if f_obj.name.endswith('.csv'):
                    df_actual = pd.read_csv(f_obj, skiprows=best_row)
                else:
                    df_actual = pd.read_excel(f_obj, skiprows=best_row)
                df_actual.columns = df_actual.columns.astype(str).str.strip()
                detail_dfs.append(df_actual)
            else:
                # 讀取 B 檔案原始排版 (鎖定 H/I 欄)
                f_obj.seek(0)
                if f_obj.name.endswith('.csv'):
                    df_b_raw = pd.read_csv(f_obj, header=None)
                else:
                    df_b_raw = pd.read_excel(f_obj, header=None)
                time_dfs.append(df_b_raw)
                
        st.write(f"📊 自動分流完成：明細組 (A類) 共 {len(detail_dfs)} 個檔案 / 時間組 (B類) 共 {len(time_dfs)} 個檔案")
        
        if len(detail_dfs) == 0:
            st.error("❌ 找不到任何銷貨明細明細檔案，請確認檔案內容。")
        elif len(time_dfs) == 0:
            st.error("❌ 找不到任何時間對照檔案（銷貨單查詢），請確認檔案內容。")
        else:
            st.success("🎯 數據庫建立成功！開始進行多天期動態解構...")
            
            # 1. 把所有上傳的明細檔合併成一個大數據庫
            full_detail_df = pd.concat(detail_dfs, ignore_index=True)
            
            # 2. 把所有時間檔的 H/I 欄位全部揉成一個大的精準時間庫
            combined_b_list = []
            for df_b_raw in time_dfs:
                if len(df_b_raw.columns) >= 9:
                    col_h = df_b_raw[7].astype(str).str.strip()
                    col_i = df_b_raw[8].astype(str).str.strip()
                    
                    h_has_time = col_h.str.contains(':', na=False).sum()
                    i_has_time = col_i.str.contains(':', na=False).sum()
                    
                    if i_has_time > h_has_time:
                        df_b_clean = pd.DataFrame({'B_客戶': df_b_raw[7], 'B_時間': df_b_raw[8]})
                    else:
                        df_b_clean = pd.DataFrame({'B_客戶': df_b_raw[8], 'B_時間': df_b_raw[7]})
                        
                    df_b_clean['客戶編號_對齊用'] = df_b_clean['B_客戶'].apply(clean_customer_id)
                    combined_b_list.append(df_b_clean)
            
            if combined_b_list:
                full_time_df = pd.concat(combined_b_list, ignore_index=True)
                full_time_df = full_time_df.dropna(subset=['客戶編號_對齊用']).drop_duplicates(subset=['客戶編號_對齊用'])
            else:
                full_time_df = pd.DataFrame(columns=['B_客戶', 'B_時間', '客戶編號_對齊用'])
            
            # 3. 關鍵突破：動態解析大數據庫中，到底包含了哪幾天的資料
            if '銷貨日' in full_detail_df.columns:
                # 把每一列的銷貨日都轉換成 YYYY-MM-DD 格式
                full_detail_df['標準日期'] = full_detail_df['銷貨日'].apply(try_parse_date)
                
                # 撈出裡面所有不重複的有效日期，排序一下
                all_detected_dates = sorted([d for d in full_detail_df['標準日期'].dropna().unique()])
                
                st.subheader("📦 跨天動態拆分結果")
                
                processed_count = 0
                for current_date in all_detected_dates:
                    # 依據目前日期切出當天的明細
                    df_day = full_detail_df[full_detail_df['標準日期'] == current_date].copy()
                    
                    # 篩選當天的豬肉相關類別
                    df_day['商品類別'] = df_day['商品類別'].astype(str).str.strip()
                    allowed_categories = ['豬肉', '豬骨', '豬冷凍']
                    df_day_filtered = df_day[df_day['商品類別'].isin(allowed_categories)].copy()
                    
                    if len(df_day_filtered) == 0:
                        continue # 如果這天沒有豬肉資料，就跳過不產出
                        
                    # 清洗當天的客戶編號
                    df_day_filtered['客戶編號_對齊用'] = df_day_filtered['客戶編號'].apply(clean_customer_id)
                    
                    # 4. 🎯 核心串接：拿當天豬肉明細去精準時間庫對對碰
                    if not full_time_df.empty:
                        merged_df = pd.merge(df_day_filtered, full_time_df, on='客戶編號_對齊用', how='left')
                        merged_df['銷貨時間'] = merged_df['B_時間']
                    else:
                        merged_df = df_day_filtered.copy()
                        merged_df['銷貨時間'] = '未對齊'
                        
                    # 5. 🎯 核心排版
                    final_columns = ['銷貨日', '銷貨時間', '數量', '客戶編號', '商品名稱']
                    final_df = merged_df[final_columns].copy()
                    
                    # 修剪格式
                    final_df['銷貨日'] = final_df['銷貨日'].apply(lambda x: try_parse_date(x) if try_parse_date(x) else str(x).split(' ')[0])
                    final_df['銷貨時間'] = final_df['銷貨時間'].astype(str).str.strip()
                    final_df['銷貨時間'] = final_df['銷貨時間'].replace({'nan': '未對齊', 'None': '未對齊'})
                    
                    # 在網頁上展開當天的折疊預覽
                    with st.expander(f"📅 檢視日期 {current_date} 的自動拆分成果 (共 {len(final_df)} 筆豬肉資料)"):
                        st.dataframe(final_df)
                        
                    # 產出當天獨立的 CSV 下載按鈕
                    csv_data = final_df.to_csv(index=False, encoding='utf-8-sig')
                    output_filename = f"{current_date}.csv"
                    
                    st.download_button(
                        label=f"📥 點我下載 {output_filename} 報表",
                        data=csv_data,
                        file_name=output_filename,
                        mime="text/csv",
                        key=f"btn_{current_date}"
                    )
                    processed_count += 1
                    
                if processed_count == 0:
                    st.warning("⏳ 讀取完畢，但數據庫中沒有篩選到任何『豬肉、豬骨、豬冷凍』的資料。")
            else:
                st.error("❌ 合併數據後找不到『銷貨日』欄位，無法進行跨天自動拆分。")
                
    except Exception as e:
        st.error(f"❌ 終極拆分處理時發生錯誤：{e}")
else:
    st.warning("⏳ 期待您的 Excel 檔案！請全選並拖曳上傳至上方區塊。")
