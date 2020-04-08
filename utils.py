import json


def set_row_and_col(row, col):
    if col % 2 == 0:
        raise ValueError("Invalid number %s" % col)
    if row >= 25:
        row = 1
        col += 2
    else:
        row += 1
    return row, col


def fill(cell_list, sheet):
    """
    将cell列表中的cell一次填充入sheet中
    :param cell_list: cell的列表
    :param sheet: 要被填充的sheet
    :return:
    """
    row, col = 0, 1
    for curr_cell in cell_list:
        while True:
            row, col = set_row_and_col(row, col)
            cell = sheet.cell(row, col)
            if cell.value is not None:
                continue
            sheet.cell(row, col, curr_cell.value)
            sheet.cell(row, col).comment = curr_cell.comment
            sheet.cell(row, col).fill = curr_cell.fill
            break


def parse_letter_to_num(alpha):
    if isinstance(alpha, int):
        return alpha
    if alpha.isdigit():
        return int(alpha)
    if alpha.isalpha():
        alpha = alpha.lower()
        return ord(alpha) - 96
    raise ValueError("requires alpha type or number!")


class Settings(object):
    @property
    def data(self):
        return self.__dict__

    def __setitem__(self, key, value):
        try:
            key_l = key.split("_")
            for k in key_l:
                if k.isupper():
                    continue
                raise ValueError("configuration must be UPPER!")
            self.__dict__["_".join(map(lambda s: s.lower(), key_l))] = value
        except Exception:
            raise ValueError("configuration must be UPPER!")


def get_settings():
    with open("conf.json", "rb") as f:
        conf = json.load(f)
        settings = Settings()
        for k, v in conf.items():
            settings[k] = v
        return settings
