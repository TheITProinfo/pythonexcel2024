import os
path=os.path.dirname(os.path.abspath(__file__)) #获取当前文件所在目录
os.chdir(path)
# print(os.getcwd())
# print(os.path.abspath(__file__))
file_list=os.listdir(path)
print(file_list)
old_book_name='product' #要替换的文件名
new_book_name='sales' #替换后的文件名
for file in file_list:
    if old_book_name in file:
        # os.rename(file,new_book_name+'.xlsx')
        # print(f'file {file} renamed to {new_book_name}.xlsx')
        if file.startswith('~$'):
            continue # 跳过临时文件
        new_file_name=file.replace(old_book_name,new_book_name) # 生成新文件名 replace 方法可以替换字符串中的字符
        old_file_path=os.path.join(path, file) # 生成旧文件路径
        print(old_file_path)
        new_file_path=os.path.join(path, new_file_name) # 生成新文件路径
        print(new_file_path)
        os.rename(old_file_path, new_file_path)
        print(f'file {file} renamed to {new_file_name}')
        

