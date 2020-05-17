from Globals import Globals
from collections import OrderedDict
from pyexcel_xlsx import save_data
from pyexcel_xlsx import get_data


# idx_englishs = ("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w","x", "y", "z")
# IDX_TO_ENGLISH_TITLE = {}
#
# def idx_to_english_title(idx):
#     index = idx
#     length = len(idx_englishs)
#     nums = []
#     while index + 1 > length:
#         num = index // length - 1
#         nums.append(idx_englishs[num])
#         index = index % length
#     nums.append(idx_englishs[index])
#     return ''.join(nums)
#
# def generate_idx_to_english_dict():
#     for i in range(500):
#         IDX_TO_ENGLISH_TITLE[i] = idx_to_english_title(i)
# generate_idx_to_english_dict()

# excel类
class Excel(object):
    def __init__(self, excel_data = None):
        self.sheets = OrderedDict()
        if excel_data:
            for key, sheet_data in excel_data.items():
                if len(sheet_data) == 0:
                    sheet = Excel.Sheet([[]])
                else:
                    sheet = Excel.Sheet(sheet_data)
                self.add_sheet(key, sheet)

    def create_sheet(self, sheet_name):
        sheet = Excel.Sheet([[]])
        self.add_sheet(sheet_name, sheet)
        return sheet

    def add_sheet(self, sheet_name, sheet):
        if hasattr(self, sheet_name):
            raise Exception(print("已经存在"+sheet_name+", 不可以重复添加！"))
        self.sheets[sheet_name] = sheet
        setattr(self, sheet_name, sheet)

    def delete_sheet(self, sheet_name):
        if hasattr(self, sheet_name):
            self.sheets[sheet_name] = None
            setattr(self, sheet_name, None)

    def Serialization(self):
        serialization_data = OrderedDict()
        for sheet_name, sheet in self.sheets.items():
            sheet_data = sheet.Serialization()
            serialization_data[sheet_name] = sheet_data
        return serialization_data

    # 内部sheet类
    class Sheet(object):
        def __init__(self, sheet_data):
            self.isLower = True
            self.title_data = sheet_data[0]  # title数据
            for i in range(len(self.title_data)):
                self.title_data[i] = str(self.title_data[i])
            sheet_data.pop(0)
            self.sheet_data = sheet_data

            self.rowNum = len(sheet_data)  # 行数
            self.colNum = len(self.title_data)  # 列数
            self.expand_col()

        #  排序
        def Sort(self, titleName):
            index = self.title_data.index(titleName)
            if index < 0:
                raise Exception("要排序的列名在sheet中不存在！")
            self.sheet_data.sort(key = lambda x:x[index])

        # 设置大小写
        def SetLower(self, b):
            self.isLower = b

        def get_repeat_title(self):
            tem_title = []
            tem_set = set()
            for title in self.title_data:
                if title in tem_set:
                    tem_title.append(title)
                else:
                    tem_set.add(title)
            return tem_title

        def append_title(self, title_name):
            if title_name in self.title_data:
                return
            self.title_data.append(title_name)
            self.colNum = len(self.title_data)
            self.expand_col()

        def append_line(self):
            self.sheet_data.append([])
            self.rowNum = len(self.sheet_data)
            self.expand_col()

        def expand_col(self):
            for row_data in self.sheet_data:
                col_offset = self.colNum - len(row_data)
                for i in range(col_offset):
                    row_data.append(None)

        def Serialization(self):
            serialization_data = []
            serialization_data.append(self.title_data)
            for row_data in self.sheet_data:
                serialization_data.append(row_data)
            return serialization_data

        def get_row_iter_data(self, row_index):
            row_offset = row_index + 1 - self.rowNum
            for i in range(row_offset):
                self.append_line()
            row_data = self.sheet_data[row_index]
            iter_data = Excel.IterData()
            iter_data.row_index = row_index
            for col_index in range(self.colNum):
                col_name = self.title_data[col_index]
                value = row_data[col_index]
                if type(value) == str and self.isLower:
                    value = value.lower()
                iter_data[col_name] = value
            return iter_data

        def set_row_iter_data(self, iter_data):
            row_index = iter_data.row_index
            row_data = self.sheet_data[row_index]
            for col_index, col_name in enumerate(self.title_data):
                row_data[col_index] = iter_data[col_name]
                iter_data.title_dict.pop(col_name)
            iter_data.title_dict.pop("row_index")
            for append_col_name, append_col_value in iter_data.title_dict.items():
                self.append_title(append_col_name)
                row_data[self.colNum - 1] = append_col_value

    # 迭代类
    class IterData(object):
        def __init__(self):
            self.title_dict = {}

        def __setattr__(self, key, value):
            if key == "title_dict":
                self.__dict__[key] = value
            else:
                self.__setitem__(key, value)

        def __setitem__(self, key, value):
            self.title_dict[key] = value
            self.__dict__[key] = value

        def __getitem__(self, key):
            return self.__dict__[key]

        def __str__(self):
            return str(self.title_dict)

# 迭代多个sheet
def iter_sheets(func, *args):
    max_row_num = 0
    for i, sheet in enumerate(args):
        max_row_num = max(max_row_num, sheet.rowNum)
        repeat_title = sheet.get_repeat_title()
        if len(repeat_title) > 0:
             raise Exception("警告：迭代时第%d个sheet有重复列：%s"%(i + 1, "，".join(repeat_title)))

    for row_index in range(max_row_num):
        func_params = []
        for sheet in args:
            iter_data = sheet.get_row_iter_data(row_index)
            func_params.append(iter_data)
        func(*func_params)
        for i, sheet in enumerate(args):
            iter_data = func_params[i]
            sheet.set_row_iter_data(iter_data)

# 从桌面文件中读取excel表
# 参数fileInDesktopName：桌面文件名
def read_excel(fileInDesktopName, isWraper = True):
    excel_data = get_data(Globals.desktop_path + fileInDesktopName)
    if isWraper:
        return Excel(excel_data)
    else:
        return excel_data

# 将excel数据写入桌面文件，文件不存在则创建
# 参数fileInDesktopName：桌面文件名
# 参数excel_object：excel数据
def write_excel(fileInDesktopName, excel_object):
    if type(excel_object) == Excel:
        excel_data = excel_object.Serialization()
    else:
        excel_data = excel_object
    save_data(Globals.desktop_path + fileInDesktopName, excel_data)