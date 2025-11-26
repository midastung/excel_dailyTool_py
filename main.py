# -*- coding: utf-8 -*-
import subprocess
import sys
from pathlib import Path
import time

# 你的所有腳本放在這個資料夾底下
BASE_DIR = Path(r"C:\Project\daily\code")

# 執行順序（依序執行）
scripts = [
    "daily_single_1.py",
    "run_dailyCopy_2.py",
    "daily_check_col_3.py",
    "daily_bundle_copy_4.py",
    "run_dailyBundleCopy_5.py",
    "dailyBundle_check_col_6.py",  # 若暫時沒有此檔，會自動略過
    "daily_unrent_7.py",
    # "daily_multiDays_8.py"
]

def run_script(script_name: str):
    """執行單一腳本，若失敗則中斷"""
    script_path = BASE_DIR / script_name
    if not script_path.exists():
        print(f"⚠️ 找不到 {script_name}，略過。")
        return
    print(f"\n🚀 執行 {script_name} ...")
    start = time.time()
    result = subprocess.run([sys.executable, str(script_path)], text=True)
    if result.returncode == 0:
        print(f"✅ {script_name} 執行完成，耗時 {time.time() - start:.2f} 秒")
    else:
        print(f"❌ {script_name} 執行失敗，中斷流程。")
        sys.exit(1)

def main():
    print("🎯 開始依序執行每日流程 ...\n")
    for s in scripts:
        run_script(s)
    print("\n🎉 所有腳本執行完畢！")

if __name__ == "__main__":
    main()
