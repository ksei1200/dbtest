from sqlalchemy import create_engine, Column, Integer, String
import openpyxl as px
import os
import time
import random

#Excell fileの読み込み
wbpp = px.load_workbook('db_sample_101119.xlsx', data_only=True)
#wbpp = xl.open_workbook('db_sample_101119.xlsx')

#最初のシートの読み込み
wspp = wbpp.active

#ヘッダー部分をキーとするための読み込み
data=[]
for row in wspp.values:
    data.append(row)

#SQLテーブル作成スクリプト生成
#table = "CREATE TABLE ews (id INTEGER PRIMARY KEY AUTOINCREMENT, "

table = 'CREATE TABLE ews (id INTEGER PRIMARY KEY AUTOINCREMENT, {0[0]}, {0[1]}, {0[2]}, {0[3]}, {0[4]},\
        {0[5]}, {0[6]},{0[7]},{0[8]} ,{0[9]},{0[10]},{0[11]},{0[12]},{0[13]},{0[14]},{0[15]},{0[16]},\
        {0[17]},{0[18]},{0[19]},{0[20]},{0[21]},{0[22]},{0[23]},{0[24]},{0[25]},{0[26]},{0[27]},{0[28]},\
        {0[29]}'.format(data[0])+")"
CREATE_EWS = table

#print(CREATE_EWS)

#既存DBファイル消去用スクリプト
#DROP_EWS2 = "DROP TABLE IF EXISTS ews"
#各Value設定キーリスト設定（不必要な文字を削除し、キー名のみが格納されたリストを作成
key = table.replace(' TEXT','').replace(' REAL','').replace(' INTEGER','').replace('CREATE TABLE ews (', '').replace(' PRIMARY KEY AUTOINCREMENT', '').replace(')', '')
key = key.replace(' ','').replace('id,','')
key_list = key.split(',')

#１行目見出しの削除
wspp.delete_rows(1)

#数値データ行数の取得
rw = wspp.max_row
#１行目のテーブル名を削除したため、最終行に”NONE”のデータが入るため、最終行を削除
wspp.delete_rows(rw)

#各Valueセット用SQLスクリプト生成
sqlstr = "INSERT INTO ews ("
for para in key_list:
    sqlstr = sqlstr + para + ','

sqlstr = sqlstr.rstrip(",")
sqlstr = sqlstr + ')' + ' VALUES(' + "?,"*(len(key_list))
sqlstr = sqlstr.strip(',')
sqlstr = sqlstr + ')'

#各Valueセット用SQLスクリプト
INSERT_EWS = sqlstr

#SQAlchemyからSQliteを指定
engine = create_engine('sqlite:///ews.db', echo=True)
"""
エクセルファイルを読み込んだデータ１〜１１の中からランダムに選択し、
それを５秒間隔で書き加える。
"""
for i in range(150):
    """
    DBにValueを挿入、DBファイルが存在しなかった場合は、ファイル・テーブル設定後にvalueを挿入、
    DBファイルが存在する場合はvalueの追加のみを行う。
    """

    if os.path.exists('ews.db') == False:
        with engine.connect() as con:
            con.execute(CREATE_EWS)
            a = random.randint(1, 11)
            con.execute(INSERT_EWS, data[a])

            # 変更したDBの中身を表示
            rows = con.execute("select * from ews;")
            for row in rows:
                print(row)
    else:
        with engine.connect() as con:
            a = random.randint(1, 11)
            con.execute(INSERT_EWS, data[a])

            #変更したDBの中身を表示
            rows = con.execute("select * from ews ORDER BY id;")
            for row in rows:
                print(row)
    i += 1
    print(i)
    time.sleep(4)

con.close()
