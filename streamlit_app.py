import streamlit as st
import pandas as pd
import io

st.title("📊 我的 Excel 自動修改工具")
st.write("請在下方上傳你的 Excel 檔案，系統會自動幫你修改數據並提供下載！")

# 讓使用者上傳檔案
uploaded_file = st.file_uploader("選擇一個 Excel 檔案", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # 讀取 Excel 檔案
        df = pd.read_excel(uploaded_file)
        
        st.success("檔案上傳成功！正在處理中...")
        
        # 顯示原本的資料給使用者看
        st.subheader("修改前的資料預覽：")
        st.dataframe(df.head())

        # --------------------------------------------------
        # ✨ 【這裡就是自動修改數據的地方】 ✨
        # 範例：如果資料裡面有一欄叫「金額」，自動乘以 2
        # 如果你想改別的，可以告訴我，我幫你改這段！
        # --------------------------------------------------
        if '金額' in df.columns:
            df['金額'] = df['金額'] * 2
            st.info("偵測到『金額』欄位，已自動將所有金額乘以 2！")
        else:
            st.warning("提示：目前範例程式只會自動修改名為『金額』的欄位。若您的 Excel 沒有這一欄，資料將保持原樣。")
        # --------------------------------------------------

        st.subheader("修改後的資料預覽：")
        st.dataframe(df.head())

        # 將修改後的資料轉回 Excel 格式供下載
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        processed_data = output.getvalue()

        # 顯示下載按鈕
        st.download_button(
            label="📥 點我下載修改後的 Excel 檔案",
            data=processed_data,
            file_name="修改後的檔案.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"處理檔案時發生錯誤: {e}")
