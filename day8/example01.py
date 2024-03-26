import os
import xlwings as xw

import pandas as pd

path=os.path.dirname(os.path.abspath(__file__))

print(path)

file_path=os.path.join(path,'salestotal.xlsx')
print(file_path)

app=xw.App(visible=False,add_book=False) # create a new instance of Excel
wb=app.books.open(file_path) # open the workbook

sheet=wb.sheets # select all the sheets

for sheet in sheet:
    print(sheet.name)
    values=sheet.range('A1').expand('table').options(pd.DataFrame, index=False, header=True).value
    print(values)
    result=values.sort_values(by='销售利润', ascending=False) # sort the values by 销售利润
    print(result)   
    sheet.range('A1').value=result # write the result to the sheet

wb.save() # save the workbook
wb.close() # close the workbook
app.quit() # quit the Excel instance



