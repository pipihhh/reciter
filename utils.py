import json
from conf import COLOR_DICT, CONF


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


__COLOR_INDEX = {}


def get_index_by_color(color):
    if len(__COLOR_INDEX):
        return __COLOR_INDEX.get(color, CONF["fill_threshold"])
    for idx, clr in COLOR_DICT.items():
        __COLOR_INDEX[clr] = idx
    return __COLOR_INDEX.get(color, CONF["fill_threshold"])


class Translation(object):
    def __init__(self, translation):
        translations = translation.split("\n")
        self._translation = {}
        for trans in translations:
            s, t = trans.split(".")
            t = t.replace("；", ";")
            self._translation[s] = t.split(";")

    @property
    def data(self):
        return self._translation

    def append(self, translation):
        for s, t in translation.data.items():
            if s not in self.data:
                self._translation[s] = t
            else:
                self._translation[s] = Translation._merge_trans(self._translation[s], t)

    @staticmethod
    def _merge_trans(l1, l2):
        trans_set = set()
        for l in l1:
            trans_set.add(l)
        for l in l2:
            trans_set.add(l)
        return list(trans_set)

    def __str__(self):
        ret = []
        for s, t in self._translation.items():
            ret.append(f"{s}.{';'.join(t)}")
        return "\n".join(ret)


def merge(trans1, trans2):
    if trans1 == "" or trans2 == "":
        return trans1 if trans2 == "" else trans2
    t1 = Translation(trans1)
    t2 = Translation(trans2)
    t1.append(t2)
    return str(t1)
