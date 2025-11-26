from dailyBundle_copy_task import copy_meal_tasks

tasks = [
    {
        "src_sheet": "貼餐包平台",
        "src_date_cell": "B2",
        "src_value_range": "S4:S14",
        "dst_sheet": "指定餐包日統計",
        "dst_date_row": 2,
        "dst_value_start_row": 4,
        "dst_value_end_row": 14,
    },
    {
        "src_sheet": "貼餐包平台",
        "src_date_cell": "B2",
        "src_value_range": "V4:V14",
        "dst_sheet": "指定餐包日統計",
        "dst_date_row": 2,
        "dst_value_start_row": 17,
        "dst_value_end_row": 27,
    },
    {
        "src_sheet": "貼餐包平台",
        "src_date_cell": "B2",
        "src_value_range": "X4:X14",
        "dst_sheet": "指定餐包日統計",
        "dst_date_row": 2,
        "dst_value_start_row": 30,
        "dst_value_end_row": 40,
    },
    # 你可以再加更多任務
]

copy_meal_tasks(tasks)
