
import os
import xlwings as xw

path=os.path.dirname(__file__)
# print(path)
file_path=os.path.join(path,'newsalesinfo')

# print(file_path)

file_list=os.listdir(file_path)  # destination file
# print(file_list)

app=xw.App(visible=False, add_book=False)
workbook=app.books.open(os.path.join(path, 'salesinfo', 'salesinfo.xlsx'))
worksheet=workbook.sheets
for i in file_list:
   if os.path.splitext(i)[1]=='.xlsx':
       workbooks=app.books.open(os.path.join(file_path, i))
       for j in worksheet:
        #    contents=j.range('A1').expand('table').value # get source  data
           contents=j.range('A1').expand('table').value # get source  data
           name=j.name
           workbooks.sheets.add(name, after=worksheet[worksheet.count-1])
           workbooks.sheets[name].range('A1').value=contents
           workbooks.save()
app.quit()
           



