# daily_single_1.py
import openpyxl
from openpyxl.utils import column_index_from_string
# 🔑 必須匯入這個特殊的類別來偵測合併儲存格
from openpyxl.cell.cell import MergedCell 
import re

def run_step(wb_src, wb_dst):
    """
    執行 Step 1: 將來源檔的 A1:K280 複製到 模板
    (已加入合併儲存格防呆機制)
    """
    try:
        # 1. 讀取來源工作表 (假設資料在第 1 頁)
        ws_src = wb_src.worksheets[0]
        
        # 2. 讀取目的工作表 (模板)
        target_sheet_name = "114年dailyTool-單日"
        
        if target_sheet_name in wb_dst.sheetnames:
            ws_dst = wb_dst[target_sheet_name]
        else:
            ws_dst = wb_dst.worksheets[0]
            print(f"警告: 找不到 '{target_sheet_name}'，寫入至 '{ws_dst.title}'")

        # 3. 執行複製 (A1:K280)
        source_range = ws_src["A1:K280"]
        
        start_row = 1
        start_col = 1
        
        for r_idx, row in enumerate(source_range):
            for c_idx, cell in enumerate(row):
                # 取得目的地的格子物件
                dst_cell = ws_dst.cell(row=start_row + r_idx, column=start_col + c_idx)
                
                # 🛑 關鍵修正：檢查目的地是否為「被合併的儲存格」
                if isinstance(dst_cell, MergedCell):
                    # 如果是合併儲存格的一部分(非首格)，它是唯讀的，必須跳過
                    continue

                # 正常寫入
                dst_cell.value = cell.value

        return True, "✅ Step 1 (daily_single) 執行成功：已複製 A1:K280 (已避開合併儲存格)"

    except Exception as e:
        return False, f"❌ Step 1 發生錯誤: {str(e)}"