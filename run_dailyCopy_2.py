# run_dailyCopy_2.py
import daily_copy_task  # 匯入剛剛改好的工人模組

def run_step(wb_src, wb_dst):
    """
    被 app.py 呼叫的主入口
    """
    
    # 定義所有複製任務 (從原檔移植)
    # 來源檔名 src_file 和 dst_prefix 在雲端版不需要了，因為 Workbook 是直接傳進來的
    # 但保留它們作為註解或參考沒關係
    tasks = [
        # --- (ALL) 系列 ---
        {
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B1", "src_value_range": "B2:B25",
            "dst_date_row": 2, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 新裝申請數
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B28", "src_value_range": "B29:B52",
            "dst_date_row": 56, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 新裝竣工數
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B55", "src_value_range": "B56:B79",
            "dst_date_row": 83, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 新裝註銷數
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B82", "src_value_range": "B83:B106",
            "dst_date_row": 110, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },l
        { # 拆機申請數
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B109", "src_value_range": "B110:B133",
            "dst_date_row": 137, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 拆機竣工數
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B136", "src_value_range": "B137:B160",
            "dst_date_row": 164, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 拆機註銷數
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B163", "src_value_range": "B164:B187",
            "dst_date_row": 191, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 異動申請數
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B190", "src_value_range": "B191:B214",
            "dst_date_row": 218, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 異動竣工數
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B217", "src_value_range": "B218:B241",
            "dst_date_row": 245, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 異動註銷數
            "src_sheet": "日統計模板", "dst_sheet": "日統計",
            "src_date_cell": "B244", "src_value_range": "B245:B268",
            "dst_date_row": 272, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        
        # --- (消客) 系列 - 來源在「無上網日統計模板」 ---
        { # 累計客戶數(消客)
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B1", "src_value_range": "B2:B25",
            "dst_date_row": 2, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 新裝申請數(消客)
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B28", "src_value_range": "B29:B52",
            "dst_date_row": 29, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 新裝竣工數(消客)
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B55", "src_value_range": "B56:B79",
            "dst_date_row": 56, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 新裝註銷數(消客)
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B82", "src_value_range": "B83:B106",
            "dst_date_row": 83, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 拆機申請數(消客)
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B109", "src_value_range": "B110:B133",
            "dst_date_row": 110, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 拆機竣工數(消客)
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B136", "src_value_range": "B137:B160",
            "dst_date_row": 137, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 拆機註銷數(消客)
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B163", "src_value_range": "B164:B187",
            "dst_date_row": 164, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 異動申請數(消客)
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B190", "src_value_range": "B191:B214",
            "dst_date_row": 191, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 異動竣工數(消客)
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B217", "src_value_range": "B218:B241",
            "dst_date_row": 218, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        { # 退租申請數(消客) - 原檔的註解寫退租申請數
            "src_sheet": "無上網日統計模板", "dst_sheet": "無上網日統計",
            "src_date_cell": "B244", "src_value_range": "B245:B268",
            "dst_date_row": 299, "dst_value_start_offset_row": 1, "dst_value_start_offset_col": 0
        },
        # 原檔最後還有一個退租竣工數的註解，但被截斷了，如果還有其他的，請依照格式補在下方
    ]

    # 呼叫工人執行
    return daily_copy_task.copy_by_mapping_openpyxl(wb_src, wb_dst, tasks)