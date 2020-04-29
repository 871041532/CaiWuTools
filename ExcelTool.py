from Globals import Globals
from collections import OrderedDict
from pyexcel_xlsx import save_data
from pyexcel_xlsx import get_data

# excel类
class Excel(object):
    def __init__(self, excel_data):
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
            self.title_data = sheet_data[0]  # title数据
            sheet_data.pop(0)
            self.sheet_data = sheet_data

            self.rowNum = len(sheet_data)  # 行数
            self.colNum = len(self.title_data)  # 列数
            self.expand_col()

        def append_title(self, title_name):
            if title_name in self.title_data:
                return
            self.title_data.append(title_name)
            self.colNum = len(self.title_data)
            self.expand_col()

        def append_line(self):
            self.sheet_data.append([])
            self.rowNum = len(self.title_data)
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
                iter_data[col_name] = value
            return iter_data

        def set_row_iter_data(self, iter_data):
            row_index = iter_data.row_index
            row_data = self.sheet_data[row_index]
            for col_index, col_name in enumerate(self.title_data):
                row_data[col_index] = iter_data[col_name]
                iter_data.__dict__.pop(col_name)
            iter_data.__dict__.pop("row_index")
            for append_col_name, append_col_value in iter_data.__dict__.items():
                self.append_title(append_col_name)
                row_data[self.colNum - 1] = append_col_value

    # 迭代类
    class IterData(object):
        def __init__(self):
            pass

        def __setitem__(self, key, value):
            self.__dict__[key] = value

        def __getitem__(self, key):
            return self.__dict__[key]

        def __str__(self):
            return str(self.__dict__)

# 迭代行并生成值
def iter_two_sheet(source_sheet, target_sheet, func):
    for row_index in range(source_sheet.rowNum):
        source_iter_data = source_sheet.get_row_iter_data(row_index)
        target_iter_data = target_sheet.get_row_iter_data(row_index)
        func(source_iter_data, target_iter_data)
        target_sheet.set_row_iter_data(target_iter_data)

# 迭代行并修改值
def iter_sheet(sheet, func):
    for row_index in range(sheet.rowNum):
        iter_data = sheet.get_row_iter_data(row_index)
        func(iter_data)
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

# excel1 = read_excel("new2.xlsx", True)
# sheet1 = excel1.页面1
# sheet2 = excel1.sheet2
# def iter_func(source, target):
#     target.name = source.name
#     target.author = source.author
#     target.args = source.args
#
# iter_row(sheet1, sheet2, iter_func)
# write_excel("new3.xlsx", excel1)