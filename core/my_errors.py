class NoFile(Exception):
    def __init__(self, file_path):
        self.file_path = file_path

    def __str__(self):
        return repr(f"无法找到文件{self.file_path}，请检查文件名称是否正确或者文件是否存在")


# TODO 使用过于繁琐，建议修改
class TheTypeError(Exception):
    def __init__(self, yes_type):  # yes_type: 原本应当使用的类型
        self.yes_type = yes_type

    def __str__(self):
        return repr(f"传入参数错误，请传入{self.yes_type}")