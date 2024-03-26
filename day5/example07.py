import os
import xlwings as xw
path=os.path.dirname(os.path.abspath(__file__)) # get current directory
os.chdir(path) # change current directory
print(path)
files_list=os.listdir(path)
print(files_list)
for i in files_list:
    if i.endswith('.xlsx'):
        print(i)
        wb=xw.Book(i) # create a workbook object
        wb.save(i)
        wb.close()
        print('done')
    else:
        print('not xlsx')

