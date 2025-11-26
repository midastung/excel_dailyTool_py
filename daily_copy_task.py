import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter
from datetime import datetime, date
import re

def get_cell_value(ws, cell_address):
    """安全讀取單一儲存格的值"""
    return ws[cell_address].value

def find_date_column(ws, row_idx, target_date):
    """
    在指定列 (row_idx) 尋找符合 target_date 的欄位索引
    """
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        cell_val = ws.cell(row=row_idx, column=col).value
        
        # 日期比對邏輯 (處理 datetime, date, string)
        if isinstance(cell_val, datetime):
            cell_val = cell_val.date()
        elif isinstance(cell_val, str):
            try:
                # 嘗試解析常見日期格式
                cell_val = datetime.strptime(cell_val, "%Y/%m/%d").date()
            except:
                pass
        
        # 假設 target_date 也是 date 物件
        if cell_val == target_date:
            return col
    return None

def copy_by_mapping_openpyxl(wb_src, wb_dst, tasks):
    """
    核心邏輯：執行 tasks 列表中的所有複製任務
    """
    logs = []
    success_count = 0
    fail_count = 0
    
    # 為了避免重複讀取 Sheet，我們在迴圈內動態獲取
    
    for idx, task in enumerate(tasks):
        task_name = f"Task {idx+1}"
        try:
            # 1. 解析來源
            src_sheet_name = task["src_sheet"]
            if src_sheet_name in wb_src.sheetnames:
                ws_src = wb_src[src_sheet_name]
            else:
                # 容錯：有些檔名寫 "日統計模板"，實際可能是 "日統計"
                # 這裡簡單處理：如果找不到就試著找包含關鍵字的
                ws_src = None
                for name in wb_src.sheetnames:
                    if src_sheet_name.replace("模板", "") in name:
                        ws_src = wb_src[name]
                        break
                if ws_src is None:
                    logs.append(f"⚠️ {task_name}: 找不到來源工作表 {src_sheet_name}")
                    fail_count += 1
                    continue

            # 2. 獲取來源日期 (通常在 B1 或特定格子)
            src_date_val = get_cell_value(ws_src, task["src_date_cell"])
            if isinstance(src_date_val, datetime):
                src_date_val = src_date_val.date()
            
            if not src_date_val:
                logs.append(f"⚠️ {task_name}: 無法從來源 {task['src_date_cell']} 讀取日期")
                fail_count += 1
                continue

            # 3. 讀取來源數值範圍 (例如 B2:B25)
            # openpyxl 回傳的是 tuple of tuples ((cell,), (cell,), ...)
            src_range_cells = ws_src[task["src_value_range"]]
            src_values = [row[0].value for row in src_range_cells] # 轉成 list

            # 4. 解析目的
            dst_sheet_name = task["dst_sheet"]
            if dst_sheet_name not in wb_dst.sheetnames:
                logs.append(f"⚠️ {task_name}: 找不到目的工作表 {dst_sheet_name}")
                fail_count += 1
                continue
            
            ws_dst = wb_dst[dst_sheet_name]
            
            # 5. 在目的檔尋找對應的日期欄位
            date_row = task["dst_date_row"]
            target_col_idx = find_date_column(ws_dst, date_row, src_date_val)
            
            if not target_col_idx:
                logs.append(f"⚠️ {task_name}: 在 {dst_sheet_name} 第 {date_row} 列找不到日期 {src_date_val}")
                fail_count += 1
                continue

            # 6. 計算寫入位置
            # 目標起始欄 = 找到的日期欄 + offset_col
            # 目標起始列 = 日期列 + offset_row
            dst_start_col = target_col_idx + task["dst_value_start_offset_col"]
            dst_start_row = date_row + task["dst_value_start_offset_row"]
            
            # 7. 寫入資料
            for i, val in enumerate(src_values):
                ws_dst.cell(row=dst_start_row + i, column=dst_start_col).value = val
                
            success_count += 1

        except Exception as e:
            logs.append(f"❌ {task_name} 錯誤: {str(e)}")
            fail_count += 1

    summary = f"Step 2 執行結束：成功 {success_count} 項，失敗 {fail_count} 項。"
    logs.append(summary)
    
    return True, logs