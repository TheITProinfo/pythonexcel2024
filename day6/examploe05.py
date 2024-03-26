import os
import xlwings as xw

path=os.path.dirname(os.path.abspath(__file__))
print(path)
file_path=os.path.join(path,'sales')
print(file_path)

app=xw.App(visible=False, add_book=False) # load excel apps

for file in os.listdir(file_path):
    if file.startswith('~$'):
        continue
    file_paths=os.path.join(file_path, file) # get absolute path of files
    wb=app.books.open(file_paths)

    wb.api.PrintOut() # call function in VB 
    
    wb.close()
    print(f'printed {file}')
app.quit()


        




