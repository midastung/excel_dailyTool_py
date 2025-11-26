import streamlit as st
import openpyxl
import io
import pandas as pd
import csv  # 引入 csv 模組
from openpyxl.utils.dataframe import dataframe_to_rows

# 匯入你的模組
import daily_single_1
import run_dailyCopy_2 

# -----------------
# 輔助函式 (超級強固版：支援 CSV 不規則欄位 + 編碼偵測)
# -----------------
def load_file(uploaded_file):
    """
    讀取 Excel 或 CSV 轉為 Workbook
    特色：
    1. 自動偵測 UTF-8 / Big5 / CP950 編碼
    2. 使用 csv 模組讀取，解決 'Expected 1 fields in line X' 的 Pandas 錯誤
    """
    if uploaded_file.name.lower().endswith('.csv'):
        # 1. 取得二進位資料
        bytes_data = uploaded_file.getvalue()
        
        # 2. 偵測編碼並解碼為字串
        text_data = None
        encoding_used = None
        
        # 嘗試 UTF-8
        try:
            text_data = bytes_data.decode('utf-8')
            encoding_used = 'utf-8'
        except UnicodeDecodeError:
            pass
            
        # 嘗試 Big5 (繁體中文常見)
        if text_data is None:
            try:
                text_data = bytes_data.decode('big5')
                encoding_used = 'big5'
            except UnicodeDecodeError:
                pass
                
        # 嘗試 CP950 (Windows 擴充繁體)
        if text_data is None:
            try:
                text_data = bytes_data.decode('cp950')
                encoding_used = 'cp950'
            except UnicodeDecodeError:
                # 真的沒招了，強制忽略錯誤讀取
                text_data = bytes_data.decode('utf-8', errors='ignore')
                encoding_used = 'ignore'

        # 3. 使用 csv 模組讀取 (容忍不規則欄位)
        f_io = io.StringIO(text_data)
        reader = csv.reader(f_io)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # 逐列寫入 Excel (不管每一列有幾個欄位，通通寫進去)
        for row in reader:
            ws.append(row)
            
        return wb
        
    else:
        # Excel 檔案直接讀取
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
                # 1. 載入檔案 (使用新版 load_file)
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