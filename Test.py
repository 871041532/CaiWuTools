import datetime
print (datetime.datetime)
value = "2018/11/06"
year, month, day = [int(x) for x in value.split('/')]
print(year)
print(month)
print(day)
# for col_name in col_names:
#     idx = 0
#     col = sheet[col_name]
#     for cell in col:
#         if idx > 0:
#             cell.number_format = "mm-dd-yy"
#         idx = idx + 1
# excel.save(file_path)