import openpyxl as px
import sqlite3

#Excell fileの読み込み
wbpp = px.load_workbook('db_sample_101119.xlsx', data_only=True)

#最初のシートの読み込み
wspp = wbpp[wbpp.sheetnames[0]]

#ヘッダー部分をキーとするための読み込み
data=[]
for row in wspp.values:
    data.append(row)

#SQLテーブル作成スクリプト生成
table = "CREATE TABLE ews1 (id INTEGER PRIMARY KEY AUTOINCREMENT, "

for header in data[0]:
    table = table + header +", "

table = table.rstrip(", ")
table = table + ")"

#SQLテーブルセット/DBファイル生成
conn = sqlite3.connect('ews1.db')
c = conn.cursor()
DROP_EWS1 = "DROP TABLE IF EXISTS ews1"
CREATE_EWS1 = table
c.execute(DROP_EWS1)
c.execute(CREATE_EWS1)

#各Value設定キー設定
table2 = table.replace(' TEXT','').replace(' REAL','').replace(' INTEGER','').replace('CREATE TABLE ews1 (', '').replace(' PRIMARY KEY AUTOINCREMENT', '').replace(')', '')
table2 = table2.replace(' ','').replace('id,','')
table2_list = table2.split(',')

#１行目見出しの削除
wspp.delete_rows(1)

#各Valueセット用SQLスクリプト生成
sqlstr = "INSERT INTO ews1 ("
for para in table2_list:
    sqlstr = sqlstr + para + ','

sqlstr = sqlstr.rstrip(",")
sqlstr = sqlstr + ')' + ' VALUES(' + "?,"*(len(table2_list))
print(sqlstr)

sqlstr = sqlstr.strip(',')
sqlstr = sqlstr + ')'

print(sqlstr)

#DB設定
for row in wspp.values:
    c.execute(sqlstr, row)
conn.commit()

#行毎のデータ表示
for resultrow in c.execute('SELECT * FROM ews1 ORDER BY id'):
    print(resultrow)

conn.close()