import os
import xlwings as xw
path=os.path.dirname(os.path.abspath(__file__)) # get current path
print(path)
file_path=os.path.join(path,'temp')
print(file_path)
file_list=os.listdir(file_path) # 
print(file_list)
old_sheet_name='Sheet1'
new_sheet_name='employeename'
app=xw.App(visible=False,add_book=False) # create a new Excel app object
for file in file_list:
    if file.startswith('~$'):
        continue
    old_file_path=os.path.join(file_path, file)
    wb=app.books.open(old_file_path) # open the workbook
    for sheet in wb.sheets:
        if sheet.name==old_sheet_name:
            sheet.name=new_sheet_name
            print(f'sheet {sheet.name} renamed to {new_sheet_name}') # rename to new sheet name
        wb.save()
        wb.close()
app.quit()
    

