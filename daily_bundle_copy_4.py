# -*- coding: utf-8 -*-
from pathlib import Path
import sys
import win32com.client as win32

# ▶ 之後所有檔案都在這個路徑底下
BASE_DIR = Path(r"C:\Project\daily")

# ───────────────────────── 基本工具 ─────────────────────────
def find_file_by_stem_or_prefix(stem_or_prefix: str) -> Path | None:
    """在 BASE_DIR 下先精準找同 stem，找不到再用 startswith 放寬（自動支援有日期/版本號的檔名）。"""
    if not BASE_DIR.exists():
        return None
    # 精準
    for p in BASE_DIR.iterdir():
        if p.is_file() and p.stem == stem_or_prefix:
            return p.resolve()
    # 放寬：startswith
    for p in BASE_DIR.iterdir():
        if p.is_file() and p.stem.startswith(stem_or_prefix):
            return p.resolve()
    return None

def open_excel():
    excel = win32.DispatchEx("Excel.Application")
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        return excel
    except Exception:
        pass
        
# ───────────────────────── 你的需求：貼 A1:K280 值 ─────────────────────────
def paste_values_keep_dest_format(
    src_prefix: str = "dailybundlemail",
    dst_stem: str = "影視業務日報表",
    dst_sheet: str = "貼餐包平台",
    src_range: str = "A1:O182",
    dst_range: str = "B2:P183",
    sheet_password: str = ""  # 若 DAY1 有保護請填入密碼
):
    src_path = find_file_by_stem_or_prefix(src_prefix)
    if not src_path:
        print(f"找不到來源檔（檔名開頭/相同為「{src_prefix}」）於：{BASE_DIR}")
        sys.exit(1)

    dst_path = find_file_by_stem_or_prefix(dst_stem)
    if not dst_path:
        print(f"找不到目的檔（檔名為/開頭為「{dst_stem}」）於：{BASE_DIR}")
        sys.exit(1)

    excel = None
    wb_src = wb_dst = None
    try:
        excel = open_excel()
        wb_src = excel.Workbooks.Open(str(src_path), ReadOnly=True)
        wb_dst = excel.Workbooks.Open(str(dst_path), ReadOnly=False)

        if wb_dst.ReadOnly:
            print("目的檔以唯讀方式開啟，請先解除唯讀或關閉其他使用者的占用。")
            sys.exit(1)

        ws_src = wb_src.Worksheets(1)  # 沒指定來源表就用第一張
        try:
            ws_dst = wb_dst.Worksheets(dst_sheet)
        except Exception:
            print(f"目的檔中找不到工作表「{dst_sheet}」")
            sys.exit(1)

        # 嘗試解除保護（沒有密碼填空字串即可）
        try:
            if ws_dst.ProtectContents or ws_dst.ProtectDrawingObjects or ws_dst.ProtectScenarios:
                ws_dst.Unprotect(Password=sheet_password)
        except Exception:
            print("目標工作表受保護且無法解除，請先取消保護或提供密碼。")
            sys.exit(1)

        src_rng = ws_src.Range(src_range)
        dst_rng = ws_dst.Range(dst_range)

        # 不用 ClearContents，直接以指定值的方式覆蓋（保留目的格式）
        try:
            dst_rng.Value = None
        except Exception:
            for c in dst_rng:
                c.Value = None

        dst_rng.Value = src_rng.Value  # 只貼值，格式不變
        wb_dst.Save()

    finally:
        if wb_src: wb_src.Close(SaveChanges=False)
        if wb_dst: wb_dst.Close(SaveChanges=False)
        if excel:  excel.Quit()

if __name__ == "__main__":
    paste_values_keep_dest_format()
