from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from core import CONF


def diff_sheets(sheet):
    words_pool = {}
    error_list = []
    for col in range(1, sheet.max_column+1, 2):
        for row in range(1, sheet.max_row+1):
            data = sheet.cell(row, col)
            if data.value is None:
                continue
            if data.value not in words_pool:
                words_pool[data.value] = f"({get_column_letter(col)}:{row})"
            else:
                error_list.append(f"{data.value}:({get_column_letter(col)}:{row}->{words_pool[data.value]})")
                data.value = None
                data.comment = None
    print(error_list)


if __name__ == '__main__':
    ws = load_workbook(CONF["pathOfExcel"])
    sheet = ws[CONF["nameOfSheets"][0]]
    diff_sheets(sheet)
    ws.save(CONF["pathOfExcel"])
