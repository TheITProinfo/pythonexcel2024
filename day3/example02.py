import os
import pandas as pd
path=r"d:\code\pythoncode\pythonexcel\day3"
os.chdir(path)
df1=pd.DataFrame({"companyname":["tom","jerry","lily"],"score":[90,95,85]})
df2=pd.DataFrame({"companyname":["Alice", "jerry", "lily"], "stockprice":[20, 180, 30]})
# df3=pd.merge(df1, df2, on="companyname")

# print(df3)
# df3=pd.merge(df1, df2, on="companyname", how="outer")
# print(df3)

## concat 

# df3=pd.concat([df1, df2], axis=0)
# print(df3)

## append
df3=df1.append(df2,ignore_index=True)
print(df3)
print(df3.index)
print(df3.columns)



df3.to_excel("test02.xlsx", sheet_name="cst100")


