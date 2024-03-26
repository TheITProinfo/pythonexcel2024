import os
import xlwings as xw

import pandas as pd

path=os.path.dirname(os.path.abspath(__file__))
print(path)
file_path=os.path.join(path,'productinfo')
print(file_path)
file_list=os.listdir(file_path)
print(file_list)
app=xw.App(visible=False, add_book=False)
for file in file_list:
    if file.startswith('~$'):
        continue
    else:
       wb=app.books.open(os.path.join(file_path, file))
       worksheet=wb.sheets['规格表']
    #    values=worksheet['A2'].expand('table').value # get all value
       values=worksheet.range('A1').options(pd.DataFrame, header=1,expand='table').value # get all value
       print(values)
       new_value=values['规格'].str.split('*', expand=True) # split the value by "*"
       values['length(mm)']=new_value[0]
       values['width(mm)']=new_value[1]
       values['height']=new_value[2]
       values.drop(columns=['规格'], inplace=True) # delete 
       worksheet['A1'].options(index=False, header=True).value=values
       worksheet.autofit()
       wb.save()
       wb.close()
app.quit()


