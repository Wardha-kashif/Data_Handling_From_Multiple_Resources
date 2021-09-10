import pandas as pd
df=pd.read_excel("sheet4.xlsx",index_col=0)
print(df)
total=50
# df.style.apply(lambda x:["background:red" if x<total else "background:green" for x in df.Percentage], axis=0)
df2=df.style.apply(lambda x:["background-color:red" if x<total else "background-color:green" for x in df.Percentage], axis=0)

df2.to_excel("sheet.xlsx", engine="openpyxl" ,index=False)
