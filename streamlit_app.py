import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime

st.set_page_config(page_title="多檔案智慧盲測過濾工具", layout="centered", page_icon="🐷")

st.title("🐷 多檔案盲測自動辨識工具")
st.write("💡 **使用說明**：您可以一次全選並上傳多個檔案（例如2天份共4個檔）。系統會完全無視 `_1`、`-1` 等系統重複檔名，直接依據**檔案內部真實數據與日期**自動進行分類配對與豬肉篩選。")

def clean_customer_id(x):
    if pd.isnull(x):
        return ""
    s = str(x).strip()
    if s.endswith('.0'):
        s = s[:-2]
    if s.isdigit():
        return str(int(s))
    return s

# 讓使用者一次上傳多個檔案
st.subheader("📁 檔案上傳區 (支援多檔案同時拖曳)")
uploaded_files = st.file_uploader("請選取並上傳您所有的 Excel 檔案", type=["xls", "xlsx", "csv"], accept_multiple_files=True)

if uploaded_files:
    try:
        st.info(f"🔍 偵測到 {len(uploaded_files)} 個檔案，智慧大腦正在讀取內部數據進行歸類...")
        
        detail_files = [] # 存放明細檔 (A檔)
        time_files = []   # 存放時間檔 (B檔)
        
        # 第一階段：完全不管檔名，純看內容欄位特徵「認人」
        for f_obj in uploaded_files:
            # 讀取前 5 列來判斷特徵
            try:
                f_obj.seek(0)
                df_preview = pd.read_excel(f_obj, header=None, nrows=5)
            except Exception:
                f_obj.seek(0)
                df_preview = pd.read_csv(f_obj, header=None, nrows=5)
            
            preview_str = str(df_preview.values)
            
            if '商品類別' in preview_str or '商品名稱' in preview_str:
                # 判定為明細檔 (A檔)
                detail_files.append(f_obj)
            else:
                # 判定為時間檔 (B檔)
                time_files.append(f_obj)
                
        st.write(f"📊 自動辨識結果：**明細檔 (A類)** 共 {len(detail_files)} 個 / **時間檔 (B類)** 共 {len(time_files)} 個")
        
        if len(detail_files) == 0:
            st.error("❌ 找不到任何包含『商品類別』的銷貨明細檔案，請確認上傳內容。")
        elif len(time_files) == 0:
            st.error("❌ 找不到任何時間對照表檔案，請確認是否成功提取 H/I 欄位檔案。")
        else:
            st.success("🎯 檔案角色分配成功！開始解析內部真實日期與進行跨日配對...")
            
            # 準備建立按「日期」分組的字典
            # 格式：{ '2026-06-17': {'detail': df, 'time_raw': df_raw}, ... }
            date_groups = {}
            
            # 解析明細檔 (A檔)，並依檔案內「銷貨日」抓出真實日期
            req_a = ['銷貨日', '數量', '客戶編號', '商品名稱', '商品類別']
            for f_obj in detail_files:
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
                    df_a = pd.read_csv(f_obj, skiprows=best_row)
                else:
                    df_a = pd.read_excel(f_obj, skiprows=best_row)
                df_a.columns = df_a.columns.astype(str).str.strip()
                
                # 找出這個檔案裡面的主要日期
                if '銷貨日' in df_a.columns and len(df_a) > 0:
                    first_date = str(df_a['銷貨日'].iloc[0]).split(' ')[0].strip()
                    # 防呆驗證是否為有效日期格式 YYYY-MM-DD 或 YYYY/MM/DD
                    if re.match(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', first_date):
                        # 統一將日期格式換成 YYYY-MM-DD
                        standard_date = first_date.replace('/', '-')
                        
                        if standard_date not in date_groups:
                            date_groups[standard_date] = {'detail': None, 'time_raw': []}
                        date_groups[standard_date]['detail'] = df_a

            # 讀取時間檔 (B檔) 原始數據
            all_b_data = []
            for f_obj in time_files:
                f_obj.seek(0)
                if f_obj.name.endswith('.csv'):
                    df_b_raw = pd.read_csv(f_obj, header=None)
                else:
                    df_b_raw = pd.read_excel(f_obj, header=None)
                all_b_data.append(df_b_raw)
                
            # 開始依照各個日期分組進行「過濾」與「跨檔時間黏合」
            st.subheader("📦 串接結果與下載區")
            
            processed_count = 0
            for current_date, group in date_groups.items():
                df_a = group['detail']
                if df_a is None:
                    continue
                    
                # 1. 篩選豬肉類別
                df_a['商品類別'] = df_a['商品類別'].astype(str).str.strip()
                allowed_categories = ['豬肉', '豬骨', '豬冷凍']
                df_a_filtered = df_a[df_a['商品類別'].isin(allowed_categories)].copy()
                
                if len(df_a_filtered) == 0:
                    st.warning(f"⚠️ 日期 {current_date} 的明細中沒有任何豬肉相關類別，跳過處理。")
                    continue
                    
                # 2. 清洗 A 檔當天的客戶編號
                df_a_filtered['客戶編號_對齊用'] = df_a_filtered['客戶編號'].apply(clean_customer_id)
                
                # 3. 智慧從所有的 B 檔中，找出能對得上當天客戶的 H/I 欄位組合
                # 建立一個當天專用的時間對照表
                combined_b_list = []
                for df_b_raw in all_b_data:
                    if len(df_b_raw.columns) >= 9: # 確保至少有到 I 欄 (索引 8)
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
                    df_b_total = pd.concat(combined_b_list, ignore_index=True)
                    df_b_final = df_b_total.dropna(subset=['客戶編號_對齊用']).drop_duplicates(subset=['客戶編號_對齊用'])
                    
                    # 4. 🎯 核心串接：跨檔案黏合正確時間
                    merged_df = pd.merge(df_a_filtered, df_b_final, on='客戶編號_對齊用', how='left')
                    merged_df['銷貨時間'] = merged_df['B_時間']
                else:
                    merged_df = df_a_filtered.copy()
                    merged_df['銷貨時間'] = '未對齊'
                    
                # 5. 🎯 核心排版
                final_columns = ['銷貨日', '銷貨時間', '數量', '客戶編號', '商品名稱']
                final_df = merged_df[final_columns].copy()
                
                # 細節格式漂亮修剪
                final_df['銷貨日'] = final_df['銷貨日'].astype(str).str.split(' ').str[0]
                final_df['銷貨時間'] = final_df['銷貨時間'].astype(str).str.strip()
                final_df['銷貨時間'] = final_df['銷貨時間'].replace({'nan': '未對齊', 'None': '未對齊'})
                
                # 顯示當天預覽
                with st.expander(f"📅 檢視日期 {current_date} 的處理成果"):
                    st.dataframe(final_df)
                    
                # 轉換並輸出 CSV
                csv_data = final_df.to_csv(index=False, encoding='utf-8-sig')
                output_filename = f"{current_date}.csv"
                
                st.download_button(
                    label=f"📥 點我下載 {output_filename} (已校正版)",
                    data=csv_data,
                    file_name=output_filename,
                    mime="text/csv",
                    key=f"btn_{current_date}"
                )
                processed_count += 1
                
            if processed_count == 0:
                st.warning("⏳ 讀取完畢，但沒有任何日期符合自動配對標準。")
                
    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤：{e}")
else:
    st.warning("⏳ 期待您的 Excel 檔案！請全選並拖曳上傳至上方區塊。")
