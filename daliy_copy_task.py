# -*- coding: utf-8 -*-
from pathlib import Path
import win32com.client as win32
from win32com.client import constants
import time
import sys

BASE_DIR = Path(r"C:\Project\daily")

def find_file(prefix: str) -> Path | None:
    """在資料夾中尋找開頭符合的檔案"""
    for p in BASE_DIR.iterdir():
        if p.is_file() and p.stem.startswith(prefix):
            return p.resolve()
    return None


def print_progress_bar(iteration, total, success, fail, length=40):
    """在終端機顯示進度條"""
    percent = iteration / total
    filled_length = int(length * percent)
    bar = '█' * filled_length + '-' * (length - filled_length)
    sys.stdout.write(
        f'\r進度: |{bar}| {iteration}/{total} ({percent*100:5.1f}%) '
        f'✅{success} ❌{fail}'
    )
    sys.stdout.flush()
    if iteration == total:
        print()  # 換行

def copy_by_mapping(tasks: list[dict]):
    """批次貼值（高效+進度條+少IO版）"""
    start_time = time.time()
    excel = win32.DispatchEx("Excel.Application")

    time.sleep(0.5)
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        excel.Calculation = constants.xlCalculationManual
    except Exception:
        pass
    try:
        excel.ScreenUpdating = False
        excel.EnableEvents = False
    except Exception:
        pass

    opened_books = {}
    used_info = {}  # 快取 UsedRange
    success = 0
    fail = 0
    total = len(tasks)

    try:
        for idx, t in enumerate(tasks, 1):
            print_progress_bar(idx - 1, total, success, fail)

            src_file = t["src_file"]
            src_sheet = t["src_sheet"]
            dst_prefix = t["dst_prefix"]
            dst_sheet = t["dst_sheet"]
            src_key_cell = t["src_key_cell"]
            src_date_cell = t["src_date_cell"]
            src_value_range = t["src_value_range"]
            dst_key_cell = t["dst_key_cell"]
            dst_date_row = t["dst_date_row"]
            dst_value_start_offset_row = t.get("dst_value_start_offset_row", 1)
            dst_value_start_offset_col = t.get("dst_value_start_offset_col", 0)

            src_path = find_file(src_file)
            dst_path = find_file(dst_prefix)

            if not src_path or not dst_path:
                fail += 1
                print_progress_bar(idx, total, success, fail)
                continue

            # 開啟或重用 Workbook
            if src_path not in opened_books:
                opened_books[src_path] = excel.Workbooks.Open(str(src_path), ReadOnly=True)
            wb_src = opened_books[src_path]

            if dst_path not in opened_books:
                opened_books[dst_path] = excel.Workbooks.Open(str(dst_path))
            wb_dst = opened_books[dst_path]

            ws_src = wb_src.Worksheets(src_sheet)
            ws_dst = wb_dst.Worksheets(dst_sheet)

            src_key = ws_src.Range(src_key_cell).Value
            src_date = ws_src.Range(src_date_cell).Value

            dst_key_col = ws_dst.Range(dst_key_cell).Column

            # 用快取避免重複讀取 UsedRange
            if ws_dst.Name not in used_info:
                used_rng = ws_dst.UsedRange
                used_rows = used_rng.Rows.Count
                used_cols = used_rng.Columns.Count
                used_info[ws_dst.Name] = (used_rows, used_cols)
            else:
                used_rows, used_cols = used_info[ws_dst.Name]

            # ⚡️ 高速批量取值 (整欄 & 整列)
            key_col_values = [v[0] for v in ws_dst.Range(
                ws_dst.Cells(1, dst_key_col),
                ws_dst.Cells(used_rows, dst_key_col)
            ).Value]

            date_row_values = list(ws_dst.Range(
                ws_dst.Cells(dst_date_row, 1),
                ws_dst.Cells(dst_date_row, used_cols)
            ).Value[0])

            try:
                row_match = key_col_values.index(src_key) + 1
            except ValueError:
                row_match = None

            try:
                col_match = date_row_values.index(src_date) + 1
            except ValueError:
                col_match = None

            if not row_match or not col_match:
                fail += 1
                print_progress_bar(idx, total, success, fail)
                continue
            
            # 📦 貼值
            src_values = ws_src.Range(src_value_range).Value
            dst_start_row = dst_date_row + dst_value_start_offset_row
            dst_start_col = col_match + dst_value_start_offset_col
            dst_end_row = dst_start_row + len(src_values) - 1

            dst_rng = ws_dst.Range(
                ws_dst.Cells(dst_start_row, dst_start_col),
                ws_dst.Cells(dst_end_row, dst_start_col)
            )

            dst_rng.Value = src_values  # 不再 Clear 以加速
            success += 1

            # 每 5 次存一次
            if idx % 5 == 0 or idx == total:
                wb_dst.Save()

            print_progress_bar(idx, total, success, fail)

        # 結尾
        print("完成: 日統計&無上網日統計 貼值")

    finally:
        for wb in list(opened_books.values()):
            wb.Close(SaveChanges=False)
        try:
            excel.Calculation = constants.xlCalculationAutomatic
            excel.ScreenUpdating = True
            excel.EnableEvents = True
        except Exception:
            pass
        excel.Quit()
