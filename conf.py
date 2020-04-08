CONF = {
    "pathOfExcel": "/Users/fudongyi/Documents/考研/生僻词总结本子.xlsx",
    "nameOfSheets": ["llyc", "ydspc", "ydcz"],
    "times": 1, "fill_threshold": 1  # 填充颜色所需要的错误次数的阈值
}

COLOR_DICT = {
    2: "DCE2F1", 3: "E3EDCD", 4: "FFF2E2", 5: "FAF9DE", 6: "FDE6E0", 7: "E9EBFE"
}

# 海天蓝  青草绿  秋叶褐  杏仁黄  胭脂红  葛巾紫

WHITE = "00FFFFFF"

RED = "E9EBFE"

SHEETS = {"difficulties", }

PARTLY_DICT = {
    # "llyc3": {
    #     "start": "a", "end": "c"
    # },
    # "pigwei1": {
    #     "start": "i", "end": "m"
    # },
    "ydspc": {
        "start": "e", "end": "k"
    }
}

MODULE_MAP = {
    "1": "RandomWordGenerator",
    "2": "PartlyWordGenerator",
    "3": "EmphasizedWordGenerator"
}
