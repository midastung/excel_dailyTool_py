import streamlit as st
import openpyxl
import io
import pandas as pd
import csv
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import date

import daily_single_1
import run_dailyCopy_2 

# -----------------
# 輔助函式
# -----------------
def load_file(uploaded_file):
    if uploaded_file.name.lower().endswith('.csv'):
        bytes_data = uploaded_file.getvalue()
        text_data = None
        
        try:
            text_data = bytes_data.decode('utf-8')
        except UnicodeDecodeError:
            pass
            
        if text_data is None:
            try:
                text_data = bytes_data.decode('big5')
            except UnicodeDecodeError:
                pass
                
        if text_data is None:
            try:
                text_data = bytes_data.decode('cp950')
            except UnicodeDecodeError:
                text_data = bytes_data.decode('utf-8', errors='ignore')

        f_io = io.StringIO(text_data)
        reader = csv.reader(f_io)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        
        for row in reader:
            ws.append(row)
            
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
        file_step1 = st.file_uploader("mailmodamount (Step 1)", type=["xlsx", "csv"], key="f1")
        
    with col2:
        st.subheader("2. 模板檔案")
        file_tpl = st.file_uploader("模板 (Template)", type=["xlsx"], key="tpl")

    # 🔑 新增：讓使用者選擇日期
    st.subheader("3. 設定")
    target_date = st.date_input("請選擇統計日期", value=date.today())

    if st.button("🚀 執行 Step 1 & 2"):
        if not file_step1 or not file_tpl:
            st.error("請上傳必要檔案！")
            return

        log_expander = st.expander("執行紀錄", expanded=True)
        
        with st.spinner("處理中..."):
            try:
                wb_src_step1 = load_file(file_step1)
                wb_dst = openpyxl.load_workbook(file_tpl)
                
                logs = []

                # --- 執行 Step 1 ---
                ok1, msg1 = daily_single_1.run_step(wb_src_step1, wb_dst)
                logs.append(msg1)
                
                # --- 執行 Step 2 ---
                # 🔑 傳入 target_date 解決無法讀取公式日期的問題
                if ok1:
                    ok2, msg2 = run_dailyCopy_2.run_step(wb_dst, wb_dst, target_date=target_date)
                    
                    if isinstance(msg2, list):
                        logs.extend(msg2)
                    else:
                        logs.append(str(msg2))
                
                with log_expander:
                    for l in logs:
                        st.write(l)

                output = io.BytesIO()
                wb_dst.save(output)
                output.seek(0)
                
                st.success("執行完成！")
                st.download_button("📥 下載整合結果", data=output, file_name=f"Result_{target_date}.xlsx")

            except Exception as e:
                st.error(f"發生錯誤: {e}")
                import traceback
                st.text(traceback.format_exc())

if __name__ == "__main__":
    main()