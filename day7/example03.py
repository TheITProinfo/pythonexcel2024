import os
import xlwings as xw

path=os.path.dirname(os.path.abspath(__file__))
print(path)
file_path=os.path.join(path,'branchinfo')
print(file_path)
file_list=os.listdir(file_path)
print(file_list)
app=xw.App(visible=False, add_book=False)
for file in file_list:
    if file.startswith('~$'):
        continue
    else:
       
        wb=app.books.open(os.path.join(file_path, file))

        print(wb.name)

        for sheet in wb.sheets:
            print(sheet.name)
            value=sheet['A2'].expand('table').value # get all value
            #print(value)
            for row, val  in  enumerate(value):
                if val== ['bag',16,65]:
                    value[row]=['bag',36,79]
                    print('cell format is adjusted')
            sheet['A2'].expand('table').value=value # set value
        wb.save()
        wb.close()
    app.quit()


            

    wb.save()
    wb.close()

app.quit()