import os
import xlwings as xw
import pandas as pd

path=os.path.dirname(__file__)

print(path)
app=xw.App(visible=False,add_book=False)
wb=app.books.open(path+'\purchase.xlsx')
sheet=wb.sheets
print(sheet)
data=[]
for i in sheet:
    values=i.range('A1').expand('table').options(pd.DataFrame).value # get all value of the cell
    filtered=values[values['采购物品']=='复印纸'] # filter the value
    if not filtered.empty:
        data.append(filtered)
new_workbook=xw.books.add()
new_sheet=new_workbook.sheets.add('cst_sheet')
new_sheet.range('A1').value=pd.concat(data,ignore_index=False)
new_workbook.save(path+'\cst_sheet.xlsx')   
wb.close()
app.quit()


