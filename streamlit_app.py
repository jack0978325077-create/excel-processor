import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="雙檔案自動串接過濾工具", layout="centered")

st.title("🐷 雙檔案自動串接過濾工具 (時間精準校正版)")
st.write("系統會自動從 A 檔過濾豬肉，並智慧讀取 B 檔精準的『銷貨時間』依『客戶編號』串接合併。")

def smart_read_excel(file_obj, required_cols):
    """
    智慧讀取 Excel：
    自動掃描前 10 列，找出包含最多目標欄位的那一列作為 Header。
    """
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
    """ 清洗客戶編號，確保兩邊格式一致才能拼對齊 """
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
file_a = st.file_uploader("請上傳 A 檔案 (含有商品類別、數量、商品名稱等)", type=["xls", "xlsx", "csv"])
file_b = st.file_uploader("請上傳 B 檔案 (含有客戶編號、準確的銷貨時間)", type=["xls", "xlsx", "csv"])

if file_a is not None and file_b is not None:
    try:
        st.info("🚀 智慧大腦正在讀取 A、B 檔案並以 B 檔時間為準進行校正...")
        
        # 定義各自需要的必要欄位 (A檔不需要管時間，B檔一定要有時間)
        req_a = ['銷貨日', '數量', '客戶編號', '商品名稱', '商品類別']
        req_b = ['客戶編號', '銷貨時間']
        
        df_a = smart_read_excel(file_a, req_a)
        df_b = smart_read_excel(file_b, req_b)
        
        missing_a = [c for c in req_a if c not in df_a.columns]
        missing_b = [c for c in req_b if c not in df_b.columns]
        
        if missing_a or missing_b:
            if missing_a: st.error(f"❌ A 檔案缺少欄位：{', '.join(missing_a)}")
            if missing_b: st.error(f"❌ B 檔案缺少欄位：{', '.join(missing_b)}")
        else:
            # 1. 處理 A 檔案：過濾豬肉類別
            df_a['商品類別'] = df_a['商品類別'].astype(str).str.strip()
            allowed_categories = ['豬肉', '豬骨', '豬冷凍']
            df_a_filtered = df_a[df_a['商品類別'].isin(allowed_categories)].copy()
            
            # 2. 強制格式對齊：把兩邊的客戶編號都轉成最乾淨的純文字
            df_a_filtered['客戶編號_對齊用'] = df_a_filtered['客戶編號'].apply(clean_customer_id)
            df_b['客戶編號_對齊用'] = df_b['客戶編號'].apply(clean_customer_id)
            
            # 3. 處理 B 檔案：只拿客戶編號跟正確的銷貨時間，並移除重複資料避免爆炸
            df_b_clean = df_b[['客戶編號_對齊用', '銷貨時間']].dropna(subset=['客戶編號_對齊用'])
            df_b_clean = df_b_clean.drop_duplicates(subset=['客戶編號_對齊用'])
            
            # 4. 🎯 核心串接：以 A 檔為主，把 B 檔中準確的 '銷貨時間' 黏過來
            merged_df = pd.merge(df_a_filtered, df_b_clean, on='客戶編號_對齊用', how='left')
            
            # 5. 🎯 核心排版：嚴格按照您要求的由左至右順序排列
            final_columns = ['銷貨日', '銷貨時間', '數量', '客戶編號', '商品名稱']
            final_df = merged_df[final_columns].copy()
            
            # 6. 細節優化（修剪日期時間格式）
            if '銷貨日' in final_df.columns:
                final_df['銷貨日'] = final_df['銷貨日'].astype(str).str.split(' ').str[0]
            if '銷貨時間' in final_df.columns:
                final_df['銷貨時間'] = final_df['銷貨時間'].astype(str).str.strip()
                final_df['銷貨時間'] = final_df['銷貨時間'].replace({'nan': '未對齊', 'None': '未對齊'})
            
            st.success("✨ 雙檔案精準校正串接成功！")
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
