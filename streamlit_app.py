import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="銷貨明細特定類別篩選器", layout="centered")

st.title("🐷 豬肉類別自動過濾工具")
st.write("目前條件設定：")
st.write("1. 系統會自動篩選，**只留下**商品類別為：`豬肉`、`豬骨`、`豬冷凍` 的數據，其餘類別自動刪除。")
st.write("2. 刪除完成後，會**自動移除商品類別欄位**，最終檔案只保留：`銷貨日`、`數量`、`客戶編號`、`商品名稱`。")
st.write("3. ✨ **新功能**：下載的檔名會自動命名為當前的 **年_月份.csv**。")

# 讓使用者上傳檔案
uploaded_file = st.file_uploader("選擇您的原始檔案 (.xls, .xlsx)", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        st.info("正在讀取並處理檔案中...")
        
        # 讀取 Excel 檔案
        df = pd.read_excel(uploaded_file)
            
        # 自動移除欄位名稱前後可能有的空格
        df.columns = df.columns.str.strip()
        
        # 檢查必要的欄位是否存在
        required_columns = ['銷貨日', '數量', '客戶編號', '商品名稱', '商品類別']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"❌ 檔案中缺少必要的欄位：{', '.join(missing_columns)}，請檢查原始檔案。")
        else:
            # 1. 先抓出需要的 5 個欄位做處理
            temp_df = df[required_columns].copy()
            
            # 清除商品類別可能含有的前後空白字元
            temp_df['商品類別'] = temp_df['商品類別'].astype(str).str.strip()
            
            # 2. 核心過濾：只留下商品類別為 豬肉、豬骨、豬冷凍 的資料
            allowed_categories = ['豬肉', '豬骨', '豬冷凍']
            filtered_df = temp_df[temp_df['商品類別'].isin(allowed_categories)].copy()
            
            # 3. 核心刪除：刪除整個「商品類別」欄位，不放進最終成果
            final_df = filtered_df.drop(columns=['商品類別'])
            
            # 優化：如果客戶編號後面有 .0 (例如 230084.0)，自動把它去掉，變成乾淨的純數字文字
            if '客戶編號' in final_df.columns:
                final_df['客戶編號'] = final_df['客戶編號'].apply(
                    lambda x: str(int(x)) if pd.notnull(x) and str(x).endswith('.0') else (str(int(x)) if isinstance(x, (int, float)) and pd.notnull(x) else x)
                )
            
            # 優化：銷貨日如果包含時間 (如 12:00:00 AM)，只保留日期部分
            if '銷貨日' in final_df.columns:
                final_df['銷貨日'] = final_df['銷貨日'].astype(str).str.split(' ').str[0]
            
            st.success(f"✨ 成功過濾！已刪除無關類別，並移除了商品類別欄位。")
            st.subheader("📋 最終 CSV 資料預覽：")
            st.dataframe(final_df)
            
            # 轉換為 CSV 格式（使用 utf-8-sig 確保用 Excel 打開時中文不會變成亂碼）
            csv_data = final_df.to_csv(index=False, encoding='utf-8-sig')
            
            # 🎯 核心功能：自動抓取今天的 年份 與 月份 作為檔名 (例如: 2026年06月.csv)
            # 註：檔名中不建議使用斜線 "/"，因為系統會誤以為是資料夾路徑導致下載失敗，所以用 "年" 和 "月" 隔開
            current_time = datetime.now()
            output_filename = current_time.strftime("%Y年%m月.csv")
            
            # 顯示下載按鈕
            st.download_button(
                label=f"📥 點我下載 {output_filename}",
                data=csv_data,
                file_name=output_filename,
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤。")
        st.error(f"錯誤訊息: {e}")
