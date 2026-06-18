import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="雙檔案自動串接過濾工具", layout="centered")

st.title("🐷 雙檔案自動串接過濾工具")
st.write("請同時提供 A、B 兩個檔案，系統會自動智慧偵測標題列並以『客戶編號』串接，過濾豬肉類別。")

def smart_read_excel(file_obj, required_cols):
    """
    智慧讀取 Excel 函數：
    自動掃描前 10 列，找出包含最多目標欄位的那一列作為 Header。
    """
    try:
        # 先不設 header 讀取前 20 列來分析
        df_preview = pd.read_excel(file_obj, header=None, nrows=20)
    except Exception:
        # 如果是 CSV
        file_obj.seek(0)
        df_preview = pd.read_csv(file_obj, header=None, nrows=20)
    
    best_row = 0
    max_matches = -1
    
    # 掃描前 10 列
    for i in range(min(10, len(df_preview))):
        row_values = [str(x).strip() for x in df_preview.iloc[i].dropna()]
        matches = sum(1 for col in required_cols if col in row_values)
        if matches > max_matches:
            max_matches = matches
            best_row = i
            
    # 回到檔案開頭重新精準讀取
    if hasattr(file_obj, 'seek'): file_obj.seek(0)
    
    if file_obj.name.endswith('.csv'):
        df = pd.read_csv(file_obj, skiprows=best_row)
    else:
        df = pd.read_excel(file_obj, skiprows=best_row)
        
    df.columns = df.columns.astype(str).str.strip()
    return df

# 讓使用者上傳兩個檔案
st.subheader("1. 上傳檔案區")
file_a = st.file_uploader("請上傳 A 檔案 (含有商品類別、數量等)", type=["xls", "xlsx", "csv"])
file_b = st.file_uploader("請上傳 B 檔案 (含有客戶編號、銷貨時間等)", type=["xls", "xlsx", "csv"])

if file_a is not None and file_b is not None:
    try:
        st.info("🚀 智慧大腦正在掃描並對齊檔案，請稍候...")
        
        # 定義各自需要的必要欄位
        req_a = ['銷貨日', '數量', '客戶編號', '商品名稱', '商品類別']
        req_b = ['客戶編號', '銷貨時間']
        
        # 使用智慧讀取器自動抓取正確的表格起點
        df_a = smart_read_excel(file_a, req_a)
        df_b = smart_read_excel(file_b, req_b)
        
        # 再次檢查欄位是否真的存在
        missing_a = [c for c in req_a if c not in df_a.columns]
        missing_b = [c for c in req_b if c not in df_b.columns]
        
        if missing_a or missing_b:
            if missing_a: st.error(f"❌ 智慧搜尋後，A 檔案仍缺少欄位：{', '.join(missing_a)}")
            if missing_b: st.error(f"❌ 智慧搜尋後，B 檔案仍缺少欄位：{', '.join(missing_b)}")
        else:
            # 1. 處理 A 檔案：過濾豬肉類別
            df_a['商品類別'] = df_a['商品類別'].astype(str).str.strip()
            allowed_categories = ['豬肉', '豬骨', '豬冷凍']
            df_a_filtered = df_a[df_a['商品類別'].isin(allowed_categories)].copy()
            
            # 清理 A 檔的客戶編號格式（轉成字串，去掉 .0）
            df_a_filtered['客戶編號'] = df_a_filtered['客戶編號'].apply(
                lambda x: str(int(x)) if pd.notnull(x) and str(x).endswith('.0') else (str(int(x)) if isinstance(x, (int, float)) and pd.notnull(x) else str(x).strip())
            )
            
            # 清理 B 檔的客戶編號格式（轉成字串，去掉 .0）
            df_b['客戶編號'] = df_b['客戶編號'].apply(
                lambda x: str(int(x)) if pd.notnull(x) and str(x).endswith('.0') else (str(int(x)) if isinstance(x, (int, float)) and pd.notnull(x) else str(x).strip())
            )
            
            # 2. 處理 B 檔案：只取客戶編號與銷貨時間，並移除重複資料避免爆炸
            df_b_clean = df_b[['客戶編號', '銷貨時間']].drop_duplicates(subset=['客戶編號'])
            
            # 3. 🎯 核心串接：兩邊都用 '客戶編號' 來串接
            merged_df = pd.merge(df_a_filtered, df_b_clean, on='客戶編號', how='left')
            
            # 4. 🎯 核心排版：嚴格按照您要求的由左至右順序排列
            final_columns = ['銷貨日', '銷貨時間', '數量', '客戶編號', '商品名稱']
            final_df = merged_df[final_columns].copy()
            
            # 優化：銷貨日如果包含時間，只保留日期部分
            if '銷貨日' in final_df.columns:
                final_df['銷貨日'] = final_df['銷貨日'].astype(str).str.split(' ').str[0]
            
            st.success("✨ 雙檔案自動對齊、串接與過濾成功！")
            st.subheader("📋 最終 CSV 資料預覽：")
            st.dataframe(final_df)
            
            # 轉換為 CSV 格式
            csv_data = final_df.to_csv(index=False, encoding='utf-8-sig')
            
            # 自動產生今天日期的檔名
            current_time = datetime.now()
            output_filename = current_time.strftime("%Y-%m-%d.csv")
            
            # 顯示下載按鈕
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
