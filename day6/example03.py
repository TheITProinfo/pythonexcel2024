import os
import xlwings as xw
path=os.path.dirname(os.path.abspath(__file__)) # get current path
print(path)
file_path=os.path.join(path,'sales') # get file path
print(file_path)

file_list=os.listdir(file_path) # get file list
print(file_list)
sheet_name='productarea'

app=xw.App(visible=False,add_book=False)
for file in file_list:
    if file.startswith('~$'):
        continue
    file_paths=os.path.join(file_path, file)
    print(file_paths)
    wb=app.books.open(file_paths)
    sheet_names=[j.name for j in wb.sheets]
    if sheet_name not in sheet_names:
       wb.sheets.add(sheet_name) # add new sheet
       print(f'sheet {sheet_name} added to {file}')
       wb.save()
       wb.close()
app.quit()
