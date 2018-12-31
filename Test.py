
file_path = "C:\\Users\\Administrator\\Desktop\\印力引入模板.xlsx"

import openpyxl
excel = openpyxl.load_workbook(file_path)
sheet = excel["凭证"]
col = sheet["C"]
cell = col[3]
print(len(col))
print(col[0].value)
print(col[0].number_format)
print(col[1].number_format)
# for col_name in col_names:
#     idx = 0
#     col = sheet[col_name]
#     for cell in col:
#         if idx > 0:
#             cell.number_format = "mm-dd-yy"
#         idx = idx + 1
# excel.save(file_path)