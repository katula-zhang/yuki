import xlrd
import xlwt
import os,sys,struct,string
DIR = os.path.abspath('.')
srcFile1 = DIR + '\\sourceFile\\' + "test1.xlsx"
srcFile2 = DIR + '\\sourceFile\\' + "test2.xlsx"
wb = xlrd.open_workbook(srcFile1)
wb_sheet = wb.sheet_by_name('Sheet1')
print(wb_sheet.col_values(1))
#sb = xlwt.open_workbook(srcFile2)
#sb_sheet = sb.get_sheet_by_name('Sheet1')