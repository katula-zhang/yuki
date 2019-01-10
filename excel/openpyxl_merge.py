#!/usr/bin/python3
from openpyxl import load_workbook
from openpyxl import workbook
import os,sys,struct,string
DIR = os.path.abspath('.')
sourceDir = DIR + '\\sourceFile\\'
objectDir = DIR = '\\objectFile\\'
tmpFile1 = '部门考核模板'
tmpFile2 = '个人考核模板'
result_Wb = load_workbook(objectDir + tmpFile2 + '.xlsx')
def copyResult(str):
        sb = load_workbook(str)
        sheet1 = sb.get_sheet_by_name('Sheet1')
        print(wb.get_sheet_names())
        count = 0
        for i in sheet['B']:
                print(i.value)
                count += 1
                sheet1["C%d" % count].value = i.value
        sb.save('test2.xlsx')

files = os.listdir(sourceDir)
tName = ['']
tCount = 'C'
for file in files :
        sub = file.split('_')
        print(sub)
        if(sub == tName[0]):#
