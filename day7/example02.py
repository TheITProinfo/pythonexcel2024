import os
import xlwings as xw

#path=os.getcwd()
path=os.path.dirname(os.path.realpath(__file__))
print(path)
file_path=os.path.join(path,'purchaseinfo')
print(file_path)
file_list=os.listdir(file_path)
print(file_list)
app=xw.App(visible=False,add_book=False)
for file in file_list:
    if file.startswith('~$'):
        continue
    else:
       wb=app.books.open(os.path.join(file_path, file))
       for sheet in wb.sheets:
            row_num=sheet.used_range.last_cell.row # get the last row number
            print(row_num)
            sheet['A2:A{}'.format(row_num)].number_format='mm/dd/yyyy'  # format the date 2018/1/31
            sheet['D2:D{}'.format(row_num)].number_format='$#,##0.00'  # format the number 2.3
            print('cell format is adjusted')

       wb.save()
       wb.close()
app.quit()

