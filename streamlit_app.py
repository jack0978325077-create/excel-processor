import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="雙檔案自動串接過濾工具", layout="centered")

st.title("🐷 雙檔案自動串接過濾工具 (精準 H/I 欄校正版)")
st.write("已特別為您修改：直接鎖定 B 檔案的 H 欄與 I 欄來精準補足實際購買時間。")

def smart_read_excel_a(file_obj, required_cols):
    """ 讀取 A 檔案：智慧搜尋標題列 """
    try:
        df_preview = pd.read_excel(file_obj, header=None, nrows=20)
    except Exception:
        file_obj.seek(0)
        df_preview = pd.read_csv(file_obj, header=None, nrows=20)
    
    best_row = 0
    max_matches = -1
    for i in range(min(10, len(df_preview))):
        row_values = [str(x).strip() for x in df_preview.iloc[i].dropna()]
        matches = sum(1 for col in required_cols if col in row_values)
        if matches > max_matches:
            max_matches = matches
            best_row = i
            
    if hasattr(file_obj, 'seek'): file_obj.seek(0)
    if file_obj.name.endswith('.csv'):
        df = pd.read_csv(file_obj, skiprows=best_row)
    else:
        df = pd.read_excel(file_obj, skiprows=best_row)
    df.columns = df.columns.astype(str).str.strip()
    return df

def clean_customer_id(x):
    if pd.isnull(x):
        return ""
    s = str(x).strip()
    if s.endswith('.0'):
        s = s[:-2]
    if s.isdigit():
        return str(int(s))
    return s

# 讓使用者上傳兩個檔案
st.subheader("1. 上傳檔案區")
file_a = st.file_uploader("請上傳 A 檔案 (含有商品類別、數量等)", type=["xls", "xlsx", "csv"])
file_b = st.file_uploader("請上傳 B 檔案 (將直接提取 H 欄與 I 欄)", type=["xls", "xlsx", "csv"])

if file_a is not None and file_b is not None:
    try:
        st.info("🚀 智慧大腦正在讀取 A 檔，並強制提取 B 檔 H 與 I 欄進行時間校正...")
        
        # 1. 讀取 A 檔案
        req_a = ['銷貨日', '數量', '客戶編號', '商品名稱', '商品類別']
        df_a = smart_read_excel_a(file_a, req_a)
        
        missing_a = [c for c in req_a if c not in df_a.columns]
        if missing_a:
            st.error(f"❌ A 檔案缺少欄位：{', '.join(missing_a)}")
        else:
            # 2. 強制不設標題讀取 B 檔案全部內容，方便我們用欄位編號去抓
            if file_b.name.endswith('.csv'):
                df_b_raw = pd.read_csv(file_b, header=None)
            else:
                df_b_raw = pd.read_excel(file_b, header=None)
            
            # Excel 的 H 欄是索引 7 (第8欄)，I 欄是索引 8 (第9欄)
            # 為了防呆，我們自動判斷這兩欄哪一個是客戶編號(通常是純數字或較短)，哪一個是時間
            col_h = df_b_raw[7].astype(str).str.strip()
            col_i = df_b_raw[8].astype(str).str.strip()
            
            # 智慧判斷：哪一欄包含冒號 ':' 或者時間格式，哪一欄就是時間
            h_has_time = col_h.str.contains(':', na=False).sum()
            i_has_time = col_i.str.contains(':', na=False).sum()
            
            if i_has_time > h_has_time:
                df_b_clean = pd.DataFrame({'B_客戶': df_b_raw[7], 'B_時間': df_b_raw[8]})
            else:
                df_b_clean = pd.DataFrame({'B_客戶': df_b_raw[8], 'B_時間': df_b_raw[7]})
            
            # 3. 處理 A 檔案：過濾豬肉類別
            df_a['商品類別'] = df_a['商品類別'].astype(str).str.strip()
            allowed_categories = ['豬肉', '豬骨', '豬冷凍']
            df_a_filtered = df_a[df_a['商品類別'].isin(allowed_categories)].copy()
            
            # 4. 強力清洗兩邊的客戶編號格式
            df_a_filtered['客戶編號_對齊用'] = df_a_filtered['客戶編號'].apply(clean_customer_id)
            df_b_clean['客戶編號_對齊用'] = df_b_clean['B_客戶'].apply(clean_customer_id)
            
            # 移除 B 檔重複資料
            df_b_final = df_b_clean.dropna(subset=['客戶編號_對齊用'])
            df_b_final = df_b_final.drop_duplicates(subset=['客戶編號_對齊用'])
            
            # 5. 🎯 核心串接：將 B 檔 H/I 欄分離出來的正確時間黏回 A 檔
            merged_df = pd.merge(df_a_filtered, df_b_final, on='客戶編號_對齊用', how='left')
            
            # 欄位重新命名為您要求的名稱
            merged_df['銷貨時間'] = merged_df['B_時間']
            
            # 6. 🎯 核心排版：由左至右
            final_columns = ['銷貨日', '銷貨時間', '數量', '客戶編號', '商品名稱']
            final_df = merged_df[final_columns].copy()
            
            # 7. 細節優化
            if '銷貨日' in final_df.columns:
                final_df['銷貨日'] = final_df['銷貨日'].astype(str).str.split(' ').str[0]
            if '銷貨時間' in final_df.columns:
                final_df['銷貨時間'] = final_df['銷貨時間'].astype(str).str.strip()
                final_df['銷貨時間'] = final_df['銷貨時間'].replace({'nan': '未對齊', 'None': '未對齊'})
            
            st.success("✨ 雙檔案精準 H/I 欄位校正對齊成功！")
            st.subheader("📋 最終 CSV 資料預覽：")
            st.dataframe(final_df)
            
            csv_data = final_df.to_csv(index=False, encoding='utf-8-sig')
            
            current_time = datetime.now()
            output_filename = current_time.strftime("%Y-%m-%d.csv")
            
            st.download_button(
                label=f"📥 點我下載 {output_filename}",
                data=csv_data,
                file_name=output_filename,
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤：{e}")
else:
    st.warning("⏳ 請同時上傳 A 檔案與 B 檔案以開始處理。")
