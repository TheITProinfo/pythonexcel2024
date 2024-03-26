import os
import xlwings as xw
import pandas as pd
path=os.path.join(os.path.dirname(__file__))
print(path)
file_path=os.path.join(path,"purchase_order.xlsx")
print(file_path)
app=xw.App(visible=False,add_book=False) # create a new instance of Excel
wb=app.books.open(file_path) # open the workbook
sheets=wb.sheets # select the first sheet
table=pd.DataFrame() # create a new DataFrame
for index,sheet in enumerate(sheets):
    values=sheet.range('A1').expand('table').options(pd.DataFrame, index=False, header=True).value
    print(values)
    data=values.reindex(columns=['采购物品','采购日期','采购数量','采购金额']) # select the columns we need
    print(data)
    table=table._append(data,ignore_index=True) # append the data to the DataFrame
print(table)
product=table[table['采购物品']=='复印纸']
print(product)
new_wb=xw.books.add() # create a new workbook
new_sheet=new_wb.sheets.add(name='复印纸') # select the first sheet
new_sheet.range('A1').value=product # copy the data to the new sheet
new_sheet.autofit() # autofit the table
new_wb.save(os.path.join(path,'purchase_order_200.xlsx')) # save the new workbook
wb.close() # close the original workbook
app.quit() # quit the Excel instance
print("done")

