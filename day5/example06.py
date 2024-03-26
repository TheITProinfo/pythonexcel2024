import os
import xlwings as xw # 

path = os.path.dirname(os.path.abspath(__file__)) # get current directory

# print(path)

os.chdir(path) # change directory to current directory
app=xw.App(visible=False,add_book=False) # create an app object
for i in range(6):
    workbook = app.books.add() # create a workbook object
    workbook.save(f'testexcel{i}.xlsx') # save the workbook

    
    workbook.close() # close the workbook
app.quit() # quit the app
print(os.listdir()) # print the list of files in the current directory
print("file saved successfully")
