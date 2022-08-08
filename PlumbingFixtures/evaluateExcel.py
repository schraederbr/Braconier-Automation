from os import getcwd
import os.path as osp
import openpyxl
import xlwings as xw



exwb = xw.Book('ToEvaluate/csm.xlsx')
exws = exwb.sheets['Plbg']

exws.range('G42').value = exws.range('G42').value
exwb.app.calculate()
exwb.save()

# wb = openpyxl.load_workbook('ToEvaluate/csm.xlsx')

# sheets = wb.sheetnames # ['Sheet1', 'Sheet2']

# for s in sheets:
#     if s != 'Plbg':
#         sheet_name = wb[s]
#         wb.remove(sheet_name)

# # your final wb with just Sheet1
# wb.save('FileCache/csm.xlsx')
# mydir = getcwd()
# import formulas
# fpath, dir_output = 'FileCache/csm.xlsx', 'ValsOnly'  # doctest: +SKIP
# xl_model = formulas.ExcelModel().loads(fpath).finish()
# xl_model.calculate()
# xl_model.write(dirpath=dir_output)
