# pip install slwings
import os
path = r"D:\code\pythoncode\pythonexcel\day2"  # set current directory    
os.chdir(path) # 
import xlwings as xw
app=xw.App(visible=True,add_book=False)
workbook=app.books.add()
workbook.save('testday200.xlsx')
print("save file successfully!!")
workbook.close()
app.quit()
