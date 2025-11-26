import streamlit as st
import openpyxl
import io
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows

# 匯入你的模組
import daily_single_1
import run_dailyCopy_2 

# -----------------
# 輔助函式
# -----------------
def load_file(uploaded_file):
    """讀取 Excel/CSV 轉為 Workbook"""
    if uploaded_file.name.lower().endswith('.csv'):
        df = pd.read_csv(uploaded_file)
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
        
        # Step 2 用的檔案 (dailybundlemail / 114年dailyTool-單日)
        
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
                # 1. 載入檔案
                wb_src_step1 = load_file(file_step1)
                wb_dst = openpyxl.load_workbook(file_tpl)
                
                logs = []

                # --- 執行 Step 1 ---
                ok1, msg1 = daily_single_1.run_step(wb_src_step1, wb_dst)
                logs.append(msg1)
                
                # --- 執行 Step 2 ---
                if ok1:
                    # 假設 Step 2 的來源就是 Step 1 剛剛修改好的 wb_dst (因為它叫 114年dailyTool-單日)
                    ok2, msg2 = run_dailyCopy_2.run_step(wb_dst, wb_dst)
                    
                    # 如果 msg2 是 list (因為我們在 daliy_copy_task 回傳了 list logs)，要展開顯示
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