import os
import xlwings as xw
path=os.path.dirname(__file__)#We love that you love Bito’s AI Code Completions!
print(path) 
file_path=os.path.join(path,'sales')
print(file_path)

sheet_name='productarea'

app=xw.App(visible=False,add_book=False) # load excel apps
for file in os.listdir(file_path):
    if file.startswith('~$'):
        continue
    file_paths=os.path.join(file_path, file)
    wb=app.books.open(file_paths)
    for sheet in wb.sheets:
        if sheet.name==sheet_name:
            print(f'sheet {sheet_name} already exists in {file}')
            sheet.delete()   # delete sheet
            # break
    wb.save
    wb.close
app.quit()
