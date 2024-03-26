import os
import xlwings as xw
import pandas as pd
path=os.path.dirname(os.path.abspath(__file__))
print(path)
file_path=os.path.join(path,'salesinfo')    
print(file_path)
flie_list=os.listdir(file_path) 
print(flie_list)
app=xw.App(visible=False, add_book=False)
for file in flie_list:
    if file.startswith('~$'):
        continue
    else:
        wb=app.books.open(os.path.join(file_path,file))
        print(wb)
        sheets=wb.sheets # select  all sheet
        print(sheets)
        for sheet in sheets:
            print(sheet.name)
            values=sheet.range('A1').expand('table').options(pd.DataFrame, index=False, header=True).value # get data from sheet
            values['销售利润']=values['销售利润'].astype(float) # convert 销售利润 to float
            print(values)
            # result=values.groupby('销售区域').sum() # group by 销售区域 and sum
            result=values.sort_values(by='销售利润', ascending=False) # sort the values by 销售利润 in descending order
           #  print(result)
            # result.to_excel(os.path.join(file_path,file[:-5]+'_'+sheet.name+'.xlsx')) # save to excel file
            sheet.range('K1').value=result # write the result to the sheet
        wb.save() # save the workbook
        wb.close()
    
app.quit()
print("done")


