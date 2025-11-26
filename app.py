import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import re
from datetime import datetime, date

# ==========================================
# 🔧 設定區 (Config)
# ==========================================
CONFIG = {
    # Step 1: 處理 mailmodamount
    "step1": {
        "src_range": "A1:K280",
        "dst_sheet": "114年dailyTool-單日", 
        "dst_start": "A1"
    },
    # Step 3: 公式修正 (針對模板)
    "step3": {
        "target_sheets": {
            "日統計": ["B4:D30", "F4:H30", "J4:L30"],
            "無上網日統計": ["B4:D30", "F4:H30", "J4:L30"]
        },
        "date_row": 3
    },
    # Step 7: 處理 mod_unrent_unfinish
    "step7": {
        "src_range": "N31:N48",
        "dst_sheet": "待拆數",
        "paste_start_row": 24,
        "target_col": 2  # 預設貼到第 2 欄 (B欄)，可依需求調整
    }
}

# ==========================================
# 核心工具函式
# ==========================================

def load_file_as_workbook(uploaded_file):
    """讀取上傳檔案，自動判斷 xlsx 或 csv 並轉為 openpyxl workbook"""
    if uploaded_file.name.lower().endswith('.csv'):
        df = pd.read_csv(uploaded_file)
        wb = openpyxl.Workbook()
        ws = wb.active
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        return wb
    else:
        return openpyxl.load_workbook(uploaded_file, data_only=True)

def copy_range_values(ws_src, ws_dst, src_range_str, dst_start_cell):
    """複製值 (Value Only)"""
    dst_col_idx = column_index_from_string(re.match(r"([A-Z]+)", dst_start_cell).group(1))
    dst_row_idx = int(re.search(r"(\d+)", dst_start_cell).group(1))
    
    src_rows = list(ws_src[src_range_str])
    
    for r_idx, row in enumerate(src_rows):
        for c_idx, cell in enumerate(row):
            ws_dst.cell(row=dst_row_idx + r_idx, column=dst_col_idx + c_idx).value = cell.value

def find_column_by_date(ws, row_idx, target_date):
    """尋找日期對應的欄位"""
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=row_idx, column=col).value
        if isinstance(val, datetime): val = val.date()
        if isinstance(val, str):
            try: val = datetime.strptime(val, "%Y/%m/%d").date()
            except: pass
        if val == target_date:
            return col
    return None

# ==========================================
# 各步驟邏輯 (Steps)
# ==========================================

# --- Step 1: 處理 mailmodamount ---
def step1_process(file_obj, wb_dst):
    st.info("執行 Step 1: 處理 mailmodamount 資料...")
    try:
        wb_src = load_file_as_workbook(file_obj)
        ws_src = wb_src.worksheets[0]
        
        # 尋找目標工作表 (如果找不到就用第一頁，或你可以指定名稱)
        target_sheet = CONFIG["step1"]["dst_sheet"]
        if target_sheet in wb_dst.sheetnames:
            ws_dst = wb_dst[target_sheet]
        else:
            ws_dst = wb_dst.worksheets[0] # 預設寫入第一頁
            
        copy_range_values(ws_src, ws_dst, CONFIG["step1"]["src_range"], CONFIG["step1"]["dst_start"])
        return True, "✅ mailmodamount 資料已複製完成"
    except Exception as e:
        return False, f"❌ Step 1 錯誤: {e}"

# --- Step 2: 處理 dailybundlemail (原本的複雜統計) ---
def step2_process(file_obj, wb_dst, target_date):
    st.info("執行 Step 2: 處理 dailybundlemail 資料 (統計分派)...")
    try:
        wb_src = load_file_as_workbook(file_obj)
        # 這裡原本是負責將 bundle 資料分派到「日統計」
        # 由於這部分邏輯較複雜且高度相依座標，這裡先保留架構
        # 你可以在此加入具體的 openpyxl 搬運邏輯
        return True, "✅ dailybundlemail 資料處理完成 (目前僅架構，需補入詳細座標)"
    except Exception as e:
        return False, f"❌ Step 2 錯誤: {e}"

# --- Step 7: 處理 mod_unrent_unfinish ---
def step7_process(file_obj, wb_dst):
    st.info("執行 Step 7: 處理 mod_unrent_unfinish 資料 (待拆數)...")
    try:
        wb_src = load_file_as_workbook(file_obj)
        ws_src = wb_src.worksheets[0]
        
        target_sheet = CONFIG["step7"]["dst_sheet"]
        if target_sheet not in wb_dst.sheetnames:
            return True, "⚠️ 無「待拆數」工作表，跳過。"

        ws_dst = wb_dst[target_sheet]
        
        # 讀取來源 N31:N48
        src_vals = [c[0].value for c in ws_src[CONFIG["step7"]["src_range"]]]
        
        # 寫入目標 (預設第2欄，可改 CONFIG)
        start_row = CONFIG["step7"]["paste_start_row"]
        col = CONFIG["step7"]["target_col"]
        
        for i, val in enumerate(src_vals):
            ws_dst.cell(row=start_row + i, column=col).value = val
            
        return True, "✅ 待拆數資料已更新"
    except Exception as e:
        return False, f"❌ Step 7 錯誤: {e}"

# --- Step 3: 公式修正 (必做) ---
def step3_fix_formulas(wb_dst, target_date):
    cfg = CONFIG["step3"]
    # 找日期欄
    check_sheet = "日統計"
    if check_sheet not in wb_dst.sheetnames:
        return True, "⚠️ 無「日統計」表，跳過公式修正。"
        
    date_col = find_column_by_date(wb_dst[check_sheet], cfg["date_row"], target_date)
    if not date_col:
        return False, f"❌ 找不到日期 {target_date}"
        
    col_letter = get_column_letter(date_col)
    pattern = re.compile(r"(日統計!)\$?[A-Z]+\$?(\d+)")
    
    count = 0
    for sheet_name, ranges in cfg["target_sheets"].items():
        if sheet_name in wb_dst.sheetnames:
            ws = wb_dst[sheet_name]
            for rng in ranges:
                # 遍歷範圍修正
                cells = ws[rng]
                if not isinstance(cells, tuple): cells = (cells,)
                for row in cells:
                    for cell in row:
                        if isinstance(cell.value, str) and "日統計!" in cell.value:
                            cell.value = pattern.sub(rf"\1{col_letter}\2", cell.value)
                            count += 1
    return True, f"✅ 公式已修正 (指向 {col_letter} 欄)"


# ==========================================
# 主程式 (UI)
# ==========================================
def main():
    st.set_page_config(page_title="影視業務日報表整合", layout="wide")
    st.title("📂 影視業務日報表整合系統")
    
    col1, col2 = st.columns([1, 1])
    
    # --- 左欄：原始資料 ---
    with col1:
        st.subheader("1. 原始資料上傳區")
        st.markdown("請一次選取以下 3 個檔案 (支援 xlsx/csv)：\n- `dailybundlemail...`\n- `mailmodamount...`\n- `mod_unrent_unfinish...`")
        uploaded_files = st.file_uploader("拖曳或選取多個檔案", accept_multiple_files=True, key="sources")
        
        # 自動分類檔案
        files_map = {}
        if uploaded_files:
            st.markdown("---")
            st.write("📂 **檔案辨識結果：**")
            for f in uploaded_files:
                fname = f.name.lower()
                if "dailybundlemail" in fname:
                    files_map["bundle"] = f
                    st.success(f"🔹 Bundle 資料: {f.name}")
                elif "mailmodamount" in fname:
                    files_map["amount"] = f
                    st.success(f"🔹 Amount 資料 (Step 1): {f.name}")
                elif "mod_unrent_unfinish" in fname:
                    files_map["unrent"] = f
                    st.success(f"🔹 待拆數資料 (Step 7): {f.name}")
                else:
                    st.warning(f"❓ 未知檔案: {f.name} (將被忽略)")

    # --- 右欄：模板 ---
    with col2:
        st.subheader("2. 模板上傳區")
        tpl_file = st.file_uploader("請上傳開頭為「影視業務日報表」的檔案", type=["xlsx"], key="template")
        if tpl_file:
            if "影視業務日報表" in tpl_file.name:
                st.success(f"✅ 已載入模板: {tpl_file.name}")
            else:
                st.warning(f"⚠️ 檔名似乎不是「影視業務日報表」，請確認是否上傳正確？({tpl_file.name})")

    # --- 下方：執行區 ---
    st.markdown("---")
    target_date = st.date_input("3. 請選擇統計日期", value=date.today())
    
    if st.button("🚀 開始整合與產出", type="primary"):
        if not tpl_file:
            st.error("❌ 請先上傳模板檔案！")
            return
            
        logs = []
        status_box = st.empty()
        
        with st.spinner("正在雲端處理資料..."):
            try:
                # 讀取模板 (這是一定要有的)
                wb_dst = openpyxl.load_workbook(tpl_file)
                
                # 依序執行各步驟
                # 1. MailModAmount (Step 1)
                if "amount" in files_map:
                    ok, msg = step1_process(files_map["amount"], wb_dst)
                    logs.append(msg)
                else:
                    logs.append("⚠️ 未上傳 mailmodamount，跳過 Step 1")

                # 2. DailyBundleMail (Step 2)
                if "bundle" in files_map:
                    ok, msg = step2_process(files_map["bundle"], wb_dst, target_date)
                    logs.append(msg)
                else:
                    logs.append("⚠️ 未上傳 dailybundlemail，跳過 Step 2")

                # 3. ModUnrent (Step 7)
                if "unrent" in files_map:
                    ok, msg = step7_process(files_map["unrent"], wb_dst)
                    logs.append(msg)
                else:
                    logs.append("⚠️ 未上傳 mod_unrent_unfinish，跳過 Step 7")

                # 4. 公式修正 (Step 3) - 只要有做任何變動最好都檢查一下公式
                ok, msg = step3_fix_formulas(wb_dst, target_date)
                logs.append(msg)

                # 顯示詳細日誌
                with st.expander("查看執行詳細報告", expanded=True):
                    for l in logs:
                        st.write(l)

                # 產出檔案
                output = io.BytesIO()
                wb_dst.save(output)
                output.seek(0)
                
                status_box.success("🎉 整合完成！請下載結果。")
                
                st.download_button(
                    label="📥 下載整合後的報表",
                    data=output,
                    file_name=f"Result_{target_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"發生系統錯誤: {str(e)}")
                import traceback
                st.text(traceback.format_exc())

if __name__ == "__main__":
    main()