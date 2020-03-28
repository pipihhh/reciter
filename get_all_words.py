import requests
from bs4 import BeautifulSoup
from bs4.element import Tag, NavigableString
from openpyxl.comments.comments import Comment
import openpyxl


SEX_SET = {"n", "vi", "vt", "art", "prep", "conj", "a", "ad", "v"}


class Word(object):
    def __init__(self, content, url, page_count):
        self._translation = {}
        self._content = content
        self._translation_cache = None
        self.page_url = url
        self.page_count = page_count

    @property
    def content(self):
        return self._content

    @property
    def translation(self):
        if self._translation_cache is None:
            t_list = []
            for sex, trans in self._translation.items():
                t_list.append(f"{sex}.{trans}")
            self._translation_cache = "\n".join(t_list)
        return self._translation_cache

    @translation.setter
    def translation(self, val):
        # trans_list = val.split(" ")
        # for trans in trans_list:
            # print(trans)
        try:
            sex, translation = val.split(".")
            self._insert(sex, translation)
        except ValueError:
            sex_set = {"n", "vi", "vt", "art", "prep", "conj", "a", "ad", "v"}
            sb_list = val.split(".")
            for i in range(len(sb_list)):
                if sb_list[i].startswith("&"):
                    sb_list[i] = sb_list[i][1:]
                for se in sex_set:
                    index = sb_list[i].find(se)
                    if index != -1 and index > 0:
                        sb_list[i] = sb_list[i][:-len(se)]
                        sb_list.insert(i+1, se)
            sex_map = {}
            former = None
            for sb in sb_list:
                if sb in sex_set:
                    sex_map[sb] = None
                else:
                    if former in sex_map and sex_map[former] is not None:
                        sex_map[former] += f";{sb}"
                    for key in sex_map:
                        if sex_map[key] is None:
                            sex_map[key] = sb
                former = sb
            for sex, translation in sex_map.items():
                self._insert(sex, translation)

    def _insert(self, sex, trans):
        if sex in self._translation and self._translation[sex] is not None:
            self._translation[sex] += f";{trans}"
        else:
            self._translation[sex] = trans
        self._translation_cache = None

    def get_dict(self):
        return {
            "content": self.content, "count": self.page_count, "url": self.page_url,
            "translation": self.translation
        }


def get_urls():
    """
    爬取此网站的所有含有词汇的url
    :return:
    """
    d = requests.get("https://kaoyan.koolearn.com/20180428/1010928.html")
    # print(d.content)
    sp4 = BeautifulSoup(d.content, "html.parser")
    tags = sp4.find_all("td", attrs={"height": "30"})
    url_list = []
    for tag in tags:
        if tag.a is not None:
            url_list.append(tag.a["href"])
    return url_list


def get_words(url, sheet):
    d = requests.get(url)
    sp4 = BeautifulSoup(d.content, "html.parser")
    tags = sp4.find_all("p", attrs={"style": "white-space: normal;"})
    words = []
    for tag in tags:
        word = words_parse(chain_tag(tag), url)
        if word is not None:
            words.append(word)
    d = {}
    for word in words:
        d[word.content] = {
            "page_count": word.page_count, "translation": word.translation,
            "url": word.page_url
        }
    col = 1
    row = 1
    i = 0
    while i < len(words):
        if row > 25:
            row = 1
            col += 2
        data = sheet.cell(row, col)
        if data.value is not None and data.comment is not None:
            row += 1
            continue
        data.value = words[i].content
        data.comment = Comment(text=words[i].translation, author="Lee Mist")
        i += 1
        row += 1
    # import json
    # with open("words.json", "w", encoding="utf-8") as f:
    #     json.dump(d, f, ensure_ascii=False)


def chain_tag(tag):
    if isinstance(tag, NavigableString):
        return tag.strip()
    else:
        ret = []
        for child in tag.children:
            if isinstance(child, NavigableString):
                ret.append(child.strip())
            else:
                ret.append(chain_tag(child))
        return "".join(ret)


def words_parse(word, url):
    word_list = word.split(" ")
    words = []
    trans_list = []
    for w in range(1, len(word_list)):
        if word_list[w].find(".") == -1:
            words.append(word_list[w])
            continue
        for se in SEX_SET:
            if word_list[w].startswith(se) and word_list[w].find(".") != -1:
                trans_list.append(word_list[w])
                break
    if len(word_list) < 3:
        return None
    word_obj = Word(" ".join(words), url, word_list[0])
    for trans in trans_list:
        word_obj.translation = trans
    return word_obj


if __name__ == '__main__':
    """
        从新东方的官网爬取5500的考研大纲词
        具体的地址看get_urls()
    """
    l = get_urls()
    wb = openpyxl.load_workbook("C:\\Users\\qq312\\Desktop\\生僻词总结本子.xlsx")
    s = wb["Sheet1"]
    for url in l:
        get_words(url, s)
    wb.save("C:\\Users\\qq312\\Desktop\\生僻词总结本子.bak2.xlsx")
    wb.close()
    # get_words("https://kaoyan.koolearn.com/20200114/1064720.html")
