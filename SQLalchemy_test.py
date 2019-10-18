from sqlalchemy import create_engine, Column, Integer, String
#from sqlalchemy.ext.declarative import declarative_base
#from sqlalchemy.orm import sessionmaker
#from sqlalchemy.orm.exc import NoResultFound

engine = create_engine('sqlite:///ews1.db', echo=True)
with engine.connect() as con:
    rows = con.execute("select * from ews1;")
    for row in rows:
        print(row)
