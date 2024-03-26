import os
import xlwings as xw
import pandas as pd
path=os.path.dirname(os.path.abspath(__file__))
print(path)
file_path=os.path.join(path,'salesbymonth.xlsx')
print(file_path)

app=xw.App(visible=False,add_book=False)
wb=app.books.open(file_path)
sheets=wb.sheets # get all sheets
for sheet in sheets:
    values=sheet.range('A1').expand('table').options(pd.DataFrame, index=False).value
    print(values)
    print(type(values))
    pivottable=pd.pivot_table(values,values='销售金额',index='销售分部',columns='销售地区',aggfunc='sum')
    print(pivottable)
    sheet.range('k1').value=pivottable
wb.save()
wb.close()
app.quit()



