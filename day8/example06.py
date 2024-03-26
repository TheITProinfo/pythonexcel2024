import os
import xlwings as xw
import pandas as pd

path=os.path.dirname(__file__)
print(path)
file_path=os.path.join(path,'salesinfo2')   
print(file_path)    

app=xw.App(visible=False, add_book=False) # create a new instance of Excel
file_list=os.listdir(file_path) # get a list of files in the directory
print(file_list)
for file in file_list:
    if file.startswith('~$'): # check if the file is an Excel file
        continue
    else:
        print(file)
        wb=app.books.open(os.path.join(file_path,file)) # open the Excel file  
        print(wb)
        worksheet=wb.sheets # get all sheets
        print(worksheet)

        for sheet in worksheet:
            values=sheet.range('A1').expand('table').options(pd.DataFrame, index=False, header=True).value # get data from sheet
            
            pivottable=pd.pivot_table(values, values=['销售金额'], index =['销售人员'], columns=['销售分部'], aggfunc='sum') # pivot the table
            print("here is the pivot table")
            print(pivottable)
            sheet.range('K1').value=pivottable # copy the data to the current sheet
        wb.save()
        wb.close()
    
app.quit() # close the Excel instance



