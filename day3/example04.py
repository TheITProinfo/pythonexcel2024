import os
import xlwings as xw
import matplotlib.pyplot as plt
figure=plt.figure()
path=r"d:\code\pythoncode\pythonexcel\day3"
os.chdir(path)


x=[1,2,3,4,5]
y=[2,4,6,8,10]
plt.plot(x,y)
app=xw.App(visible=False,add_book=False)
workbook=app.books.add()
worksheet=workbook.sheets.add('cst100')
worksheet.pictures.add(figure, name='picture1', update=True, left=200, top=0)
workbook.save('test100.xlsx')
print("file svaed suessfully!")
workbook.close()
app.quit()

