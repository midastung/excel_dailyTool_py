import streamlit as st
import openpyxl
import io
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows

# 匯入你的模組
import daily_single_1
import run_dailyCopy_2 

# -----------------
# 輔助函式 (已修正編碼問題)
# -----------------
def load_file(uploaded_file):
    """讀取 Excel/CSV 轉為 Workbook，並處理中文編碼"""
    if uploaded_file.name.lower().endswith('.csv'):
        try:
            # 1. 先嘗試 UTF-8
            df = pd.read_csv(uploaded_file, encoding='utf-8')
        except UnicodeDecodeError:
            # 2. 失敗則嘗試 Big5 (台灣系統常見)
            uploaded_file.seek(0) # 歸零指標
            try:
                df = pd.read_csv(uploaded_file, encoding='big5')
            except UnicodeDecodeError:
                # 3. 再失敗嘗試 CP950
                uploaded_file.seek(0)
                df = pd.read_csv(uploaded_file, encoding='cp950')

        wb = openpyxl.Workbook()
        ws = wb.active
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        return wb
    else:
        return openpyxl.load_workbook(uploaded_file, data_only=True)

# -----------------
# 主介面
# -----------------
def main():
    st.set_page_config(page_title="Excel 整合系統", layout="wide")
    st.title("📂 模組化 Excel 整合系統")

    # 介面配置
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. 來源檔案")
        # Step 1 用的檔案
        file_step1 = st.file_uploader("mailmodamount (Step 1)", type=["xlsx", "csv"], key="f1")
        
    with col2:
        st.subheader("2. 模板檔案")
        file_tpl = st.file_uploader("模板 (Template)", type=["xlsx"], key="tpl")

    if st.button("🚀 執行 Step 1 & 2"):
        if not file_step1 or not file_tpl:
            st.error("請上傳必要檔案！")
            return

        log_expander = st.expander("執行紀錄", expanded=True)
        
        with st.spinner("處理中..."):
            try:
                # 1. 載入檔案 (現在支援 Big5 CSV 了)
                wb_src_step1 = load_file(file_step1)
                wb_dst = openpyxl.load_workbook(file_tpl)
                
                logs = []

                # --- 執行 Step 1 ---
                ok1, msg1 = daily_single_1.run_step(wb_src_step1, wb_dst)
                logs.append(msg1)
                
                # --- 執行 Step 2 ---
                # Step 2 使用 Step 1 處理完的 wb_dst 作為來源與目的
                if ok1:
                    ok2, msg2 = run_dailyCopy_2.run_step(wb_dst, wb_dst)
                    
                    if isinstance(msg2, list):
                        logs.extend(msg2)
                    else:
                        logs.append(str(msg2))
                
                # 顯示紀錄
                with log_expander:
                    for l in logs:
                        st.write(l)

                # 下載結果
                output = io.BytesIO()
                wb_dst.save(output)
                output.seek(0)
                
                st.success("執行完成！")
                st.download_button("📥 下載整合結果", data=output, file_name="Result_Step1_2.xlsx")

            except Exception as e:
                st.error(f"發生錯誤: {e}")
                import traceback
                st.text(traceback.format_exc())

if __name__ == "__main__":
    main()