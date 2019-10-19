from sqlalchemy import create_engine, Column, Integer, String
import openpyxl as px
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
from sqlalchemy.orm.exc import NoResultFound

#Excell fileの読み込み
wbpp = px.load_workbook('db_sample_101119.xlsx', data_only=True)

#最初のシートの読み込み
wspp = wbpp[wbpp.sheetnames[0]]

#ヘッダー部分をキーとするための読み込み
data=[]
for row in wspp.values:
    data.append(row)

#SQLテーブル作成スクリプト生成
table = "CREATE TABLE ews2 (id INTEGER PRIMARY KEY AUTOINCREMENT, "

for header in data[0]:
    table = table + header +", "

table = table.rstrip(", ")
table = table + ")"
CREATE_EWS2 = table
DROP_EWS2 = "DROP TABLE IF EXISTS ews2"

#各Value設定キー設定
table2 = table.replace(' TEXT','').replace(' REAL','').replace(' INTEGER','').replace('CREATE TABLE ews2 (', '').replace(' PRIMARY KEY AUTOINCREMENT', '').replace(')', '')
table2 = table2.replace(' ','').replace('id,','')
table2_list = table2.split(',')

#１行目見出しの削除
wspp.delete_rows(1)

#各Valueセット用SQLスクリプト生成
sqlstr = "INSERT INTO ews2 ("
for para in table2_list:
    sqlstr = sqlstr + para + ','

sqlstr = sqlstr.rstrip(",")
sqlstr = sqlstr + ')' + ' VALUES(' + "?,"*(len(table2_list))
print(sqlstr)

sqlstr = sqlstr.strip(',')
sqlstr = sqlstr + ')'
print(sqlstr)

INSERT_EWS2 = sqlstr

engine = create_engine('sqlite:///ews2.db', echo=True)

with engine.connect() as con:
    # Drop文を実行
    con.execute(DROP_EWS2)
    # テーブルの作成
    con.execute(CREATE_EWS2)
    # Insert文を実行する
    #con.execute(INSERT_EWS2)
    for row in wspp.values:
        con.execute(INSERT_EWS2, row)

    rows = con.execute("select * from ews2;")
    for row in rows:
        print(row)

con.close()
