import os
import openpyxl
path = r"D:\code\pythoncode\pythonexcel"
os.chdir(path)
wb = openpyxl.load_workbook("test001.xlsx")
print(wb.sheetnames) # print all sheet names
sheet = wb['cstsheet1']
print(sheet) # print current sheet name