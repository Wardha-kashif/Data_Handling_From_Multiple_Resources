import mysql.connector
import pandas as pd

mydb = mysql.connector.connect(
host = 'localhost',
user='root',
password='warda170199',
port=3306,
database='coursedb'
)

exl_to_csv= pd.read_csv (r'C:\Users\warda.kashif\PycharmProjects\DataHandling\Student.csv.txt')
exl_to_csv.to_excel ('Student.xlsx')

cur=pd.read_sql('Select * from course', con = mydb)

df_sql_data = pd.DataFrame(cur)
df_sql_data.head()
df_sql_data.to_excel("output.xlsx")

pd.set_option('display.expand_frame_repr', False,)


mydb.close()
# Data Coming From Database
print("Data Coming From Database")
df1=cur
print(df1)


print("\n\nData Change from csv to excel\n")
df2=pd.read_excel("Student.xlsx",index_col=0)
print(df2)

print("\n\nData which is already in excel format\n")
df3=pd.read_excel("Record.xlsx")
print(df3)


df4 = df2.merge(df3, on="STUDENTID", how='outer')
print(df4)

df5=df4.merge(df1, on="courseId", how='outer')
print(df5)

df5.to_excel("final_output.xlsx")
