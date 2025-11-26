# -*- coding: utf-8 -*-
from pathlib import Path
import win32com.client as win32
from win32com.client import constants
import time
import sys

BASE_DIR = Path(r"C:\Project\daily")

# ========== 共用函式 ==========
def find_file(prefix: str) -> Path | None:
    for p in BASE_DIR.iterdir():
        if p.is_file() and p.stem.startswith(prefix):
            return p.resolve()
    return None

def num_to_excel_col(n: int) -> str:
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

def clean_str(val):
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        if float(val).is_integer():
            return str(int(val))
        else:
            return str(val)
    return str(val).strip()

def print_progress_bar(iteration, total, success, fail, length=40):
    percent = iteration / total
    filled_length = int(length * percent)
    bar = '█' * filled_length + '-' * (length - filled_length)
    sys.stdout.write(
        f'\r進度: |{bar}| {iteration}/{total} ({percent*100:5.1f}%) ✅{success} ❌{fail}'
    )
    sys.stdout.flush()
    if iteration == total:
        print()

# ========== 子功能：轉公式為值 ==========
def _range_to_values(excel_app, rng):
    """把 rng 內的公式強制轉成純值"""
    try:
        vals = rng.Value2
        rng.Value2 = vals
        return True
    except Exception:
        pass
    try:
        rng.Copy()
        rng.PasteSpecial(Paste=constants.xlPasteValues)
        excel_app.CutCopyMode = False
        return True
    except Exception:
        pass
    try:
        vals = rng.Value2
        if isinstance(vals, tuple):
            rows = len(vals)
            cols = len(vals[0]) if rows else 0
            for r in range(rows):
                for c in range(cols):
                    try:
                        rng.Cells(r + 1, c + 1).Value2 = vals[r][c]
                    except Exception:
                        pass
        else:
            rng.Value2 = vals
        return True
    except Exception:
        return False

# ========== 子功能：兩張表轉純值 ==========
def update_reward_package_format(excel, wb, src_val):
    src_str = clean_str(src_val)

    targets = [
        {
            "sheet": "獎勵餐包新裝竣工數",
            "rows": [(4, 21), (26, 43)],
            "match_rows": [3, 25],
        },
        {
            "sheet": "各指定餐包分營運處日統計",
            "rows": [(5, 443)],
            "match_rows": [3],
        },
    ]

    for t in targets:
        sheet_name = t["sheet"]
        ws = wb.Worksheets(sheet_name)
        used_cols = ws.UsedRange.Columns.Count
        
        # 🧩 限制「各指定餐包分營運處日統計」的搜尋範圍到 ND 欄
        if sheet_name == "各指定餐包分營運處日統計":
            used_cols = min(used_cols, 368)  # ND 是第 368 欄
        
        # 🧩 限制「獎勵餐包新裝竣工數」的搜尋範圍到 GE 欄
        if sheet_name == "獎勵餐包新裝竣工數":
            used_cols = min(used_cols, 188)  # GE 是第 188 欄

        try:
            ws.Unprotect()
        except Exception:
            pass
        try:
            excel.CalculateFullRebuild()
        except Exception:
            pass

        header_values = []
        for row_idx in t["match_rows"]:
            vals = [clean_str(v) for v in ws.Range(
                ws.Cells(row_idx, 1), ws.Cells(row_idx, used_cols)
            ).Value[0]]
            header_values.extend(vals)

        col_match_list = [i + 1 for i, v in enumerate(header_values) if v == src_str]
        if not col_match_list:
            print(f"❌ 找不到與「{src_str}」相符的欄位，略過 {sheet_name}")
            continue

        for col_idx in col_match_list:
            col_letter = num_to_excel_col(col_idx)

            for start_row, end_row in t["rows"]:
                try:
                    rng = ws.Range(f"{col_letter}{start_row}:{col_letter}{end_row}")
                    ok = _range_to_values(excel, rng)
                 
                except Exception as e:
                    print(f"⚠️ 無法處理 {col_letter}{start_row}:{col_letter}{end_row}：{e}")

    wb.Save()

# ========== 主功能：餐包貼值任務 ==========
def copy_meal_tasks(tasks: list[dict]):
    start_time = time.time()
    dst_path = find_file("影視業務日報表")
    if not dst_path:
        print("❌ 找不到『影視業務日報表』檔案")
        return

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        excel.Calculation = constants.xlCalculationManual
        excel.ScreenUpdating = False
        excel.EnableEvents = False
    except Exception:
        pass

    success = 0
    fail = 0
    total = len(tasks)
    src_val_for_next = None

    try:
        wb = excel.Workbooks.Open(str(dst_path))
        for idx, t in enumerate(tasks, 1):
            print_progress_bar(idx - 1, total, success, fail)
            src_sheet = t["src_sheet"]
            src_date_cell = t["src_date_cell"]
            src_value_range = t["src_value_range"]
            dst_sheet = t["dst_sheet"]
            dst_date_row = t["dst_date_row"]
            dst_value_start_row = t["dst_value_start_row"]
            dst_value_end_row = t["dst_value_end_row"]

            try:
                ws_src = wb.Worksheets(src_sheet)
                ws_dst = wb.Worksheets(dst_sheet)
                src_raw_value = ws_src.Range(src_date_cell).Value
                src_str = clean_str(src_raw_value)
                src_val_for_next = src_str
                src_values = ws_src.Range(src_value_range).Value

                used_cols = ws_dst.UsedRange.Columns.Count
                date_row_values = list(ws_dst.Range(
                    ws_dst.Cells(dst_date_row, 1),
                    ws_dst.Cells(dst_date_row, used_cols)
                ).Value[0])

                col_match = None
                for i, val in enumerate(date_row_values, start=1):
                    if clean_str(val) == src_str:
                        col_match = i
                        break

                if not col_match:
                    print(f"\n❌ 找不到相符數值「{src_str}」，略過 {src_sheet}")
                    fail += 1
                    print_progress_bar(idx, total, success, fail)
                    continue

                col_letter = num_to_excel_col(col_match)
                dst_rng = ws_dst.Range(f"{col_letter}{dst_value_start_row}:{col_letter}{dst_value_end_row}")
                dst_rng.Value = src_values

                success += 1
                if idx % 3 == 0 or idx == total:
                    wb.Save()

            except Exception as e:
                print(f"\n⚠️ 任務 {src_sheet} 發生錯誤：{e}")
                fail += 1

            print_progress_bar(idx, total, success, fail)

        wb.Save()

        # ✅ 自動執行兩張表的轉值流程
        if src_val_for_next:
            print(f"\n🧩 轉換公式為值中（依來源值 {src_val_for_next}）...")
            update_reward_package_format(excel, wb, src_val_for_next)

    finally:
        try:
            wb.Close(SaveChanges=False)
            excel.Calculation = constants.xlCalculationAutomatic
            excel.ScreenUpdating = True
            excel.EnableEvents = True
        except Exception:
            pass
        excel.Quit()
