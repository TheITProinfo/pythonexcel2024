import os

import numpy as np
import pandas as pd
path=r"d:\code\pythoncode\pythonexcel\day2"
os.chdir(path)
a=np.arange(12).reshape(3,4)
b=pd.DataFrame(a,index=[1,2,3],columns=['A','B','C','D'])
print(b)
b.to_excel('test.xlsx', sheet_name='Sheet1')
print(b.to_excel('test.xlsx', sheet_name='cst100'))

# 读取excel文件
import pandas as pd
df = pd.read_excel('test.xlsx', sheet_name='cst100')
print(df)

# 读取csv文件