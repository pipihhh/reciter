import json
import playsound


class Word(object):
    def __init__(self, content, translation):
        self.content = content
        self.translation = {}
        self._part_of_speech = {"vi", "vt", "adj", "adv", "n", "conj", "prep", "v"}

    def _insert(self, speech, translation):
        if speech in self.translation:
            self.translation[speech] = self.translation[speech].append(translation)
        else:
            self.translation[speech] = [translation, ]

    def add_translation_by_string(self, translation):
        t_list = translation.split(".")
        if len(t_list) > 2 or len(t_list) <= 1:
            raise ValueError("错误的格式")
        self._insert(t_list[0], t_list[1])

    def add_translation_by_dict(self, speech, translation):
        if speech not in self._part_of_speech:
            raise ValueError("错误的speech")
        self._insert(speech, translation)


def over_all(dic, offset):
    for a, b in dic.items():
        if isinstance(b, dict):
            print(f"{a}:")
            print("展开")
            over_all(b, offset + " ")
        else:
            print(f"{offset}{a}:{b}")


def init_words(d, dic):
    pass


def play(f_obj):
    playsound.playsound("http://media.shanbay.com/audio/us/hello.mp3")


if __name__ == '__main__':
    # dic = {}
    # with open("KaoYanluan_1.json", "r", encoding="utf-8") as f:
    #     word = f.readline()
    #     d = json.loads(word)
    #     over_all(d, "")
    import requests
    resp = requests.get("http://media.shanbay.com/audio/us/hello.mp3")
    play(resp.content)
