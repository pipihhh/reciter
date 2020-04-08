from command_handler import CommandHandler
import generators


if __name__ == '__main__':
    """
        核心的背单词的文件 只需要运行此文件即可
        y为认识此单词 n为不认识此单词 eof为退出
        error为rollback上一个单词入单词池
        excel需要提取的sheet在108行的sheets_set来指定 可以指定多个set
        在背诵过程中error过或n过的单词在程序正常结束后会被记入此文件的difficulties的sheet中
        需要人为创建这个名为difficulties的sheet或者自行指定一个名称
    """
    modules_name = list(map(lambda s: s.capitalize(), input("module_name:").split("_")))
    module = getattr(generators, f"{''.join(modules_name)}Generator")
    wordsDb = module.init_all_words()
    handler = CommandHandler(wordsDb)
    handler.loop()
