import generators
from conf import MODULE_MAP
from tkinter.font import Font
from tkinter import Tk, Label, Button, Entry, messagebox, Frame


class ClientApp(object):
    def __init__(self):
        self._app = Tk()
        self._app.title("阿皮bdc")
        self._app.geometry("900x430")
        self._frame = Frame(bg="#FAF9DE")
        word_font = Font(family="Monaco", size=30)
        translation_font = Font(family="STKaiti", size=30)
        self._word_label = Label(self._frame, text="", bg="#EAEAEF", width="200", font=word_font)
        self._translation_label = Label(self._frame, text="", bg="#C7EDCC", width="200", height="10",
                                        font=translation_font)
        self._text = Entry(self._frame)
        self._button = Button(self._frame, text="conformity", command=self._load)
        self._words_db = None
        self._current_word = None

    def run(self):
        self._frame.pack(side="top")
        self._word_label.focus_set()
        self._translation_label.focus_set()
        self._word_label.pack()
        self._translation_label.pack()
        self._text.pack()
        self._button.pack()
        self._bind()
        self._app.mainloop()

    def _bind(self):
        self._app.bind("y", self._pass_handler)
        self._app.bind("WM_DELETE_WINDOW", self._app.iconify)
        self._app.bind("s", self._show_handler)
        self._app.bind("n", self._forget_handler)
        self._app.bind("<Escape>", self._destroy)
        self._app.bind("<Control-u>", self._reinitialize)

    def _pass_handler(self, event):
        if self._message_valid():
            self._words_db.pop(self._current_word)
            if self._words_db.is_empty():
                self._reinitialize()
                return
            self._current_word = self._words_db.get_random()
            self._word_label["text"] = f"{self._current_word}({len(self._words_db)})"
            self._translation_label["text"] = ""

    def _show_handler(self, e):
        translation = self._words_db.get_translation(self._current_word)
        self._translation_label["text"] = translation

    def _forget_handler(self, e):
        if self._message_valid():
            self._words_db.add_log(self._current_word)
            self._current_word = self._words_db.get_random()
            self._word_label["text"] = f"{self._current_word}({len(self._words_db)})"
            self._translation_label["text"] = ""

    def _message_valid(self):
        try:
            if self._words_db.get_translation(self._current_word) != self._translation_label["text"]:
                messagebox.showwarning(title="Hi!", message="宁还没康答案就开始选上了?")
                return False
        except AttributeError:
            return False
        return True

    def _get_module(self, m_name: str):
        if m_name.isdigit():
            return getattr(generators, MODULE_MAP[m_name])
        else:
            module_names = list(map(lambda s: s.capitalize(), m_name.split("_")))
            module = getattr(generators, f"{''.join(module_names)}Generator")
            return module

    def _load(self):
        module_name = self._text.get()
        try:
            module = self._get_module(module_name)
            self._words_db = module.init_all_words()
            self._button.pack_forget()
            self._text.pack_forget()
            self._current_word = self._words_db.get_random()
            self._word_label["text"] = f"{self._current_word}({len(self._words_db)})"
        except Exception as e:
            messagebox.showwarning(title="Hi!", message=str(e))

    def _destroy(self, e):
        try:
            self._words_db.finish()
        except AttributeError:
            pass
        self._app.destroy()

    def _reinitialize(self, e=None):
        self._words_db.finish()
        self._button = Button(self._frame, text="conformity", command=self._load)
        self._text = Entry(self._frame)
        self._text.pack()
        self._button.pack()
        self._word_label["text"] = ""
        self._translation_label["text"] = ""
        self._words_db = None
        self._current_word = None


if __name__ == '__main__':
    ClientApp().run()
