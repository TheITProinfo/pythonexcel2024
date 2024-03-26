import os
import xlwings as xw
import pandas as pd

path=os.path.join(os.path.dirname(__file__),'purchase_order100.xlsx')
print(path)
app=xw.App(visible=False,add_book=False)
wb=app.books.open(path)
sheets=wb.sheets # get all sheets
for sheet in sheets:
    print(sheet.name)
    values=sheet.range('A1').expand('table') #    get all values in the sheet
    data=values.options(pd.DataFrame, index=False, header=True).value # convert the values to a new DataFrame
    print(data) 
    sums=data['采购金额'].sum() # calculate the sum of each column\
    print(sums)
    column=values.value[0].index('采购金额')+1# get the index of the column '采购金额'
    print(column)
    row=values.last_cell.row # get the last row
    print(row)
    sheet.range(row+1,column).value=sums # add the sum to the last row of the column '采购金额'
wb.save()
wb.close()
app.quit()
print("done")
