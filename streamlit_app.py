import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="銷貨明細欄位篩選器", layout="centered")

st.title("📊 銷貨明細特定欄位篩選工具")
st.write("目前設定：系統會自動幫您保留 **銷貨日、數量、客戶編號、商品名稱、商品類別** 這五個欄位，並轉為 CSV 供您下載。")

# 讓使用者上傳檔案
uploaded_file = st.file_uploader("選擇您的原始檔案 (.xls, .xlsx, .csv)", type=["xls", "xlsx", "csv"])

if uploaded_file is not None:
    try:
        st.info("正在讀取並處理檔案中...")
        
        # 根據副檔名自動選擇讀取方式
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        # 自動移除欄位名稱前後可能有的空格
        df.columns = df.columns.str.strip()
        
        # 🎯 您指定要保留的 5 個欄位
        target_columns = ['銷貨日', '數量', '客戶編號', '商品名稱', '商品類別']
        
        # 檢查檔案中實際存在哪些目標欄位
        available_columns = [col for col in target_columns if col in df.columns]
        missing_columns = [col for col in target_columns if col not in df.columns]
        
        if missing_columns:
            st.warning(f"提示：上傳的檔案中好像缺少以下欄位：{', '.join(missing_columns)}")
            
        if not available_columns:
            st.error("❌ 找不到任何符合的欄位，請確認您的檔案欄位名稱是否正確。")
        else:
            # 只篩選出您要的這幾個欄位
            final_df = df[available_columns]
            
            # 優化：如果客戶編號後面有 .0 (例如 230084.0)，自動把它去掉，變成乾淨的文字
            if '客戶編號' in final_df.columns:
                final_df['客戶編號'] = final_df['客戶編號'].apply(
                    lambda x: str(int(x)) if pd.notnull(x) and str(x).endswith('.0') else (str(int(x)) if isinstance(x, (int, float)) and pd.notnull(x) else x)
                )
            
            st.success("✨ 欄位篩選成功！")
            st.subheader("📋 轉換後的資料預覽：")
            st.dataframe(final_df)
            
            # 轉換為 CSV 格式（使用 utf-8-sig 確保用 Excel 打開時中文不會變成亂碼）
            csv_data = final_df.to_csv(index=False, encoding='utf-8-sig')
            
            # 自動產生新檔名
            original_name = uploaded_file.name.rsplit('.', 1)[0]
            output_filename = f"{original_name}_精簡版.csv"
            
            # 顯示下載按鈕
            st.download_button(
                label="📥 點我下載最終 CSV 檔案",
                data=csv_data,
                file_name=output_filename,
                mime="text/csv"
            )

    except Exception as e:
        st.error(f"❌ 處理檔案時發生錯誤。")
        st.error(f"錯誤訊息: {e}")
