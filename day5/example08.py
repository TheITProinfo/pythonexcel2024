import os
import xlwings as xw

path=os.path.dirname(os.path.abspath(__file__)) #获取当前文件所在目录
os.chdir(path) #切换当前目录
app=xw.App(visible=False,add_book=False) #创建excel应用
workbook=xw.Book('examplesalestotal.xlsx') #打开excel文件
worksheets=workbook.sheets #get all sheets   
for i in range(len(worksheets)):
    print("previous sheet name")
    print(worksheets[i].name)
    worksheets[i].name=worksheets[i].name.replace('销售分部','sales') #将所有sheet名称中的'销售分部 '替换为'sales'cls
    print("after name:")
    print(worksheets[i].name)
workbook.save("newexamplesalestotal.xlsx") #保存文件
workbook.close() #关闭文件
app.quit() #退出excel   
