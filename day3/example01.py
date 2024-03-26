import os
import pandas as pd

# short cut for comment ctrl+/

path=r"d:\code\pythoncode\pythonexcel\day3"
os.chdir(path) #改变当前工作目录
data=pd.DataFrame([[1,2,3],[4,5,6],[7,8,9]],index=["r1","r2","r3"],columns=["C1","C2","C3"])
print(data)
# data.to_csv("test.csv", index=False)

# a=data['C1']
# print(a)
# print(type(a))

# a1=data[['C1','C3']]
# print(a1)

# b1=data[1:3]
# print(b1)
# b2=data.iloc[1:3]
# print(b2)
# b3=data.loc['r2':'r3']
# print(b3)

# a=data[['C1','C3']][0:2]
# print(a)
# b=data.loc[['r1','r2'],['C1','C3']]
# print(b)
# compare data
# a=data[data['C1']>1]
# print(a)
# b=data[(data['C1']>1) & (data['C2']==5)]
# print(b)

## sorting

a=data.sort_values(by='C2',ascending=False)
print(a)

##  calculator

# data['C4']=data['C1']+data['C2']
# print(data)
# data['C5']=data['C1']*data['C2']
# print(data)
# data['C6']=data['C1']/data['C2']
# print(data)
# data['C7']=data['C1']-data['C2']

## delete data

# data.drop(columns='C3',inplace=True)
# print(data)













