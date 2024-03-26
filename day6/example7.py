import os
import xlwings as xw
app = xw.App(visible = False, add_book = False) 
file_path = 'd:\\table\\销售表'   
file_list = os.listdir(file_path)  
workbook = app.books.open('d:\\code\\pythoncode\\pythonexcel\\dsy\\信息表.xlsx')  
worksheet = workbook.sheets
for i in file_list:  
    if os.path.splitext(i)[1] == '.xlsx':  
        workbooks = app.books.open(file_path + '\\' + i)  
        for j in worksheet:  
            contents = j.range('A1').expand('table').value  
            name = j.name  
            workbooks.sheets.add(name = name, after = workbooks.sheets[-1])  
            workbooks.sheets[name].range('A1').value = contents  
        workbooks.save()      
app.quit()