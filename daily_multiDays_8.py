# -*- coding: utf-8 -*-
from pathlib import Path
import win32com.client as win32
import sys
import datetime

# =========================================================
# 基本設定
# =========================================================
BASE_DIR = Path(r"C:\Project\daily")


# =========================================================
# 進度條
# =========================================================
def print_progress_bar(current, total, prefix="", length=40):
    """顯示進度條動畫（█ + 百分比）"""
    percent = current / total if total else 1
    filled = int(length * percent)
    bar = "█" * filled + "-" * (length - filled)
    sys.stdout.write(f"\r{prefix} |{bar}| {percent*100:5.1f}%")
    sys.stdout.flush()
    if current == total:
        sys.stdout.write("\n")


# =========================================================
# 找檔案
# =========================================================
def find_file(prefix: str):
    for p in BASE_DIR.iterdir():
        if p.is_file() and p.name.startswith(prefix) and p.suffix.lower() == ".xlsx":
            return str(p.resolve())
    return None


# =========================================================
# 數字 → Excel 欄位字母
# =========================================================
def col_letter(col_num):
    if col_num is None:
        return None
    try:
        col_num = int(col_num)
    except:
        return None
    if col_num <= 0:
        return None

    s = ""
    while col_num > 0:
        col_num, r = divmod(col_num - 1, 26)
        s = chr(65 + r) + s
    return s


# =========================================================
# 在特定 row 中找日期所在欄
# =========================================================
def find_date_col_in_row(ws, date_value, target_row):
    last_col = ws.Cells(target_row, ws.Columns.Count).End(-4159).Column  # xlToLeft
    for c in range(1, last_col + 1):
        try:
            if ws.Cells(target_row, c).Value == date_value:
                return c
        except:
            continue
    return None


# =========================================================
# 找出欄位中「原本是公式」且為合併儲存格左上角的列
# =========================================================
def get_rows_to_update(ws, col_letter: str) -> list[int]:
    """
    高速找出指定欄位中「原本是公式」的列。
    優化方式：一次性讀取整欄的公式和合併儲存格資訊，在記憶體中處理。
    """
    col_index = ws.Range(col_letter + "1").Column
    last_row = ws.Cells(ws.Rows.Count, col_index).End(-4162).Row  # xlUp
    if last_row <= 1:
        return []

    # ⚡️ 一次性讀取整欄的公式和合併儲存格資訊
    formulas = ws.Range(f"{col_letter}1:{col_letter}{last_row}").Formula
    
    # 獲取合併儲存格的資訊
    merge_info = []
    for r in range(1, last_row + 1):
        cell = ws.Cells(r, col_index)
        merge_info.append(cell.MergeCells and cell.MergeArea.Cells(1, 1).Address == cell.Address)

    rows = []
    # 在 Python 記憶體中高速處理
    for r_idx in range(last_row):
        row_num = r_idx + 1
        formula = formulas[r_idx][0]
        is_merge_top_left = merge_info[r_idx]
        is_formula = isinstance(formula, str) and formula.startswith("=")

        if is_formula and (not ws.Cells(row_num, col_index).MergeCells or is_merge_top_left):
            rows.append(row_num)

    return rows


# =========================================================
# 套用 SUM(前日:今日) 公式（高速 + 有進度條）
# =========================================================
def apply_sum_fast(ws, col_letter_target, rows, before_letter, now_letter):

    if before_letter is None or now_letter is None:
        print(f"⚠ 日期欄位錯誤 → 跳過 {ws.Name}")
        return

    col_index = ws.Range(col_letter_target + "1").Column
    total = len(rows)
    updated = 0

    print(f"  → [{ws.Name}] {col_letter_target} 欄位開始更新公式...")

    for idx, r in enumerate(rows, start=1):

        print_progress_bar(idx, total, prefix=f"    更新列 {r}")

        try:
            cell = ws.Cells(r, col_index)
        except:
            continue

        try:
            if cell.MergeCells:
                merge = cell.MergeArea
                if merge.Row != r or merge.Column != col_index:
                    continue
                cell = merge.Cells(1, 1)

            formula = f"=SUM({before_letter}{r}:{now_letter}{r})"
            cell.Formula = formula
            updated += 1

        except:
            continue


# =========================================================
# 主流程
# =========================================================
def run():

    # -----------------------------------------------------
    # 1. 找檔案
    # -----------------------------------------------------
    file_path = find_file("影視業務日報表")
    if not file_path:
        print("❌ 找不到檔案")
        return

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(file_path)

    # -----------------------------------------------------
    # 2. 讀 J2 / M1
    # -----------------------------------------------------
    ws_summary = wb.Worksheets("摘要表")
    ws_mod = wb.Worksheets("MOD_個+企")

    j2_value = ws_summary.Range("J2").Value
    target_date = ws_mod.Range("M1").Value

    print(f"✔ J2={j2_value}, 日期(M1)={target_date}")

    if j2_value == 1:
        print("➡ J2=1 → 不需運行，結束")
        wb.Close(SaveChanges=False)
        excel.Quit()
        return
    
    # -----------------------------------------------------
    # 3. 啟動條件：昨日 == M1 才允許執行
    # -----------------------------------------------------
    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)

    # 將 M1 轉成日期物件
    try:
        if isinstance(target_date, float):
            base = datetime.date(1899, 12, 30)
            m1_date = base + datetime.timedelta(days=int(target_date))
        elif isinstance(target_date, datetime.datetime):
            m1_date = target_date.date()
        elif isinstance(target_date, datetime.date):
            m1_date = target_date
        else:
            m1_date = datetime.datetime.strptime(str(target_date), "%Y/%m/%d").date()
    except:
        print("❌ M1 日期格式錯誤，無法解析")
        wb.Close(SaveChanges=False)
        excel.Quit()
        return

    if m1_date != yesterday:
        print(f"➡ 條件未達成 → 昨日 {yesterday} ≠ M1 {m1_date} → 不執行")
        wb.Close(SaveChanges=False)
        excel.Quit()
        return

    print(f"✔ 日期條件成立：昨日 {yesterday} == M1 {m1_date} → 開始執行任務")

    # -----------------------------------------------------
    # 4. 各表的日期行
    # -----------------------------------------------------
    date_rows = {
        "日統計": 29,
        "無上網日統計": 29,
        "各指定餐包分營運處日統計": 4,
        "各指定餐包分營運處日統計 (權重數)": 4,
    }

    # -----------------------------------------------------
    # 5. 每表要處理的欄位
    # -----------------------------------------------------
    target_columns = {
        "日統計": "NH",
        "無上網日統計": "NH",
        "各指定餐包分營運處日統計": "NG",
        "各指定餐包分營運處日統計 (權重數)": "NG",
    }

    # -----------------------------------------------------
    # 6. 逐工作表處理
    # -----------------------------------------------------
    total_sheets = len(target_columns)
   

    for sheet_name, col_letter_target in target_columns.items():

        ws = wb.Worksheets(sheet_name)
        print(f"\n--- 處理工作表：{sheet_name} ---")

        if sheet_name not in date_rows:
            print(f"⚠ {sheet_name} 無日期列設定 → 跳過")
            continue

        target_row = date_rows[sheet_name]

        # 找出日期欄
        day_col = find_date_col_in_row(ws, target_date, target_row)
        if day_col is None:
            print(f"❌ {sheet_name} 找不到日期 {target_date}")
            continue

        day_row_letter = col_letter(day_col)
        day_minus = j2_value - 1
        day_before = day_col - day_minus
        day_before_letter = col_letter(day_before)

        # 找出「原本是公式」的列（高速）
        rows_to_update = get_rows_to_update(ws, col_letter_target)

        # 套用公式（高速 + 動態進度條）
        apply_sum_fast(ws, col_letter_target, rows_to_update,
                       day_before_letter, day_row_letter)

    print("\n✨ 全部工作表處理完畢！")

    wb.Close(SaveChanges=True)
    excel.Quit()


# =========================================================
# RUN
# =========================================================
if __name__ == "__main__":
    run()
