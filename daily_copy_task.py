# daliy_copy_task.py
import openpyxl
from datetime import datetime, date

def get_cell_value(ws, cell_address):
    """安全讀取單一儲存格的值"""
    try:
        return ws[cell_address].value
    except:
        return None

def find_date_column(ws, row_idx, target_date):
    """
    在指定列 (row_idx) 尋找符合 target_date 的欄位索引
    """
    max_col = ws.max_column
    # 為了效能，通常日期不會跑太遠，或是可以設定一個範圍
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
    執行 tasks 列表中的所有複製任務
    參數:
      wb_src: 來源 Excel 物件 (通常是 Step 1 處理完的結果)
      wb_dst: 目的 Excel 物件 (模板)
      tasks: 任務列表 (包含 src_key_cell 等所有參數)
    """
    logs = []
    success_count = 0
    fail_count = 0
    
    for idx, task in enumerate(tasks):
        # 取得任務名稱 (用 Key Cell 的值來當作 Log 名稱比較好辨識)
        task_label = f"Task {idx+1}"
        
        try:
            # 1. 解析來源 Sheet
            src_sheet_name = task["src_sheet"]
            # 容錯處理：移除 "模板" 兩字來比對 (因為有時候來源被改名了)
            ws_src = None
            if src_sheet_name in wb_src.sheetnames:
                ws_src = wb_src[src_sheet_name]
            else:
                # 嘗試模糊比對
                for name in wb_src.sheetnames:
                    if name in src_sheet_name or src_sheet_name.replace("模板", "") in name:
                        ws_src = wb_src[name]
                        break
            
            if ws_src is None:
                logs.append(f"⚠️ {task_label}: 找不到來源工作表 '{src_sheet_name}'")
                fail_count += 1
                continue

            # 2. 獲取來源日期 (依據 src_date_cell)
            # 雖然我們有 input date，但原始邏輯是從 Excel 格子讀取，我們保持這邏輯
            src_date_val = get_cell_value(ws_src, task["src_date_cell"])
            if isinstance(src_date_val, datetime):
                src_date_val = src_date_val.date()
            
            if not src_date_val:
                logs.append(f"⚠️ {task_label}: 無法從 {task['src_date_cell']} 讀取日期")
                fail_count += 1
                continue

            # 3. 讀取來源資料 (src_value_range)
            src_range_cells = ws_src[task["src_value_range"]]
            # 轉成平面 List [數值1, 數值2...]
            src_values = [row[0].value for row in src_range_cells]

            # (選用) 驗證 Key Cell，雖然不影響複製，但可用於確認是否對齊
            # src_key_val = get_cell_value(ws_src, task["src_key_cell"])

            # 4. 解析目的 Sheet
            dst_sheet_name = task["dst_sheet"]
            if dst_sheet_name not in wb_dst.sheetnames:
                logs.append(f"⚠️ {task_label}: 找不到目的工作表 '{dst_sheet_name}'")
                fail_count += 1
                continue
            
            ws_dst = wb_dst[dst_sheet_name]
            
            # 5. 在目的檔尋找對應的日期欄位
            # 依據 dst_date_row 這一列去橫向掃描
            date_row = task["dst_date_row"]
            target_col_idx = find_date_column(ws_dst, date_row, src_date_val)
            
            if not target_col_idx:
                logs.append(f"⚠️ {task_label}: 在 '{dst_sheet_name}' 第 {date_row} 列找不到日期 {src_date_val}")
                fail_count += 1
                continue

            # 6. 計算寫入位置
            # 目標欄位 = 找到的日期欄 + 偏移量
            # 目標列數 = 日期列 + 偏移量
            dst_start_col = target_col_idx + task["dst_value_start_offset_col"]
            dst_start_row = date_row + task["dst_value_start_offset_row"]
            
            # 7. 執行寫入
            for i, val in enumerate(src_values):
                ws_dst.cell(row=dst_start_row + i, column=dst_start_col).value = val
                
            success_count += 1
            # logs.append(f"✅ {task_label} 完成") # 若不想 Log 太多可註解

        except Exception as e:
            logs.append(f"❌ {task_label} 發生錯誤: {str(e)}")
            fail_count += 1

    summary = f"✅ Step 2 彙總：成功 {success_count} 項，失敗 {fail_count} 項。"
    logs.append(summary)
    
    # 回傳 (是否成功, Log列表)
    return True, logs