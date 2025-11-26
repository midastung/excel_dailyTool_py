import openpyxl
from openpyxl.utils import column_index_from_string
import re

def run_step(wb_src, wb_dst):
    try:
        # 1. 讀取來源工作表 (假設資料在第 1 頁)
        ws_src = wb_src.worksheets[0]
        
        # 2. 讀取目的工作表 (模板)
        # 你的原始程式碼中，目的 Sheet 名稱通常與檔名相關
        # 這裡預設找 "template_daily"
        target_sheet_name = "template_daily"
        
        if target_sheet_name in wb_dst.sheetnames:
            ws_dst = wb_dst[target_sheet_name]
        else:
            # 如果找不到指定名稱，就用第一頁，並回報警告
            ws_dst = wb_dst.worksheets[0]
            print(f"警告: 找不到 '{target_sheet_name}'，寫入至 '{ws_dst.title}'")

        # 3. 執行複製 (A1:K280) -> 貼到 (A1)
        # 來源範圍
        source_range = ws_src["A1:K280"]
        
        # 起始寫入位置 (A1)
        start_row = 1
        start_col = 1
        
        # 開始搬運 (只複製值 Value)
        for r_idx, row in enumerate(source_range):
            for c_idx, cell in enumerate(row):
                # 寫入目標格
                ws_dst.cell(row=start_row + r_idx, column=start_col + c_idx).value = cell.value

        return True, "✅ Step 1 (daily_single) 執行成功：已複製 A1:K280"

    except Exception as e:
        return False, f"❌ Step 1 發生錯誤: {str(e)}"