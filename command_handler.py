from time import time
from importlib import import_module


class CommandHandler(object):
    def __init__(self, db):
        self._lazy_class_map = {}
        self._db = db

    def parse(self, command, words_db):
        if command in self._lazy_class_map:
            self._lazy_class_map[command].process(command, words_db)
        else:
            try:
                module = import_module(f"{command}.{command.capitalize()}")
                self._lazy_class_map[command] = module()
                self._lazy_class_map[command].process(command, words_db)
            except ImportError:
                print("此命令不可用")
            except Exception as e:
                print(str(e))

    def loop(self):
        while not self._db.is_empty():
            word = self._db.get_random()
            print(f"{word}({len(self._db)})")
            t = time()
            _ = input()
            translation = self._db.get_translation(word)
            print("答案:")
            print(translation)
            print("\n")
            ans = input()
            if ans.upper() == "Y":
                self._db.set_work_seconds(word, time()-t)
                self._db.pop(word)
            elif ans.upper() == "EOF":
                break
            elif ans.lower() == "n":
                self._db.add_log(word)
            else:
                self.parse(ans, self._db)
        self._db.finish()
