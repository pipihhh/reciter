from importlib import import_module


class CommandHandler(object):
    def __init__(self):
        self._lazy_class_map = {}

    def handle(self, command, words_db):
        if command in self._lazy_class_map:
            self._lazy_class_map[command].process(command, words_db)
        else:
            try:
                module = import_module(f"{command}.{command.capitalize()}")
                self._lazy_class_map[command] = module()
                self._lazy_class_map[command].process(command, words_db)
            except ImportError:
                print("此命令不可用")
