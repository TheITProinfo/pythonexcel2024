# pip install xlwings
import os
import os
import xlwings as xw
path=r"d:\code\pythoncode\pythonexcel\day2"
os.chdir(path)
app=xw.App(visible=False,add_book=False)
wb=app.books.add()
worksheet=wb.sheets.add("product")
worksheet.range("A1").value="prduct NO."
worksheet.range("B1").value="product name"
worksheet.range("C1").value="price"
worksheet.range("A2").value=1
worksheet.range("B2").value="apple"
worksheet.range("C2").value=900
wb.save("product.xlsx")
print("file saved successfully")
wb.close()
app.quit()

