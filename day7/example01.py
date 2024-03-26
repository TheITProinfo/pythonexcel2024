
import os
import xlwings as xw

path=os.path.dirname(__file__)
file_path=os.path.join(path,'salesinfo')  # get path of files
print(file_path)
files=os.listdir(file_path)
print(files)

app=xw.App(visible=False, add_book=False) # load apps 
for file in files:
    if file.startswith('~$'):
        continue
    else:
        wb=app.books.open(os.path.join(file_path, file))
        for sheet in wb.sheets:  # get all sheets in the file
            print(sheet.name)
            value=sheet.range('A1').expand('table')  # get all cells in the sheet
            value.column_width=120             # 20  characters
            value.row_height=20               # 20  points
            print('cell format is adjusted')
        wb.save()
        wb.close()
app.quit()


