from xlrd import open_workbook

from sqlalchemy import Table, Column, Integer, String, MetaData

from sqlalchemy.orm import mapper, sessionmaker
from sqlalchemy import create_engine


db_uri = 'sqlite:///db.sqlite'
engine = create_engine(db_uri)

metadata = MetaData(engine)
metadata.drop_all()

class table_class:
    pass

def get_file_path():
    print("Please enter tne xlsx file path\n" \
          "[]>",end = " ")
    file_path = input()
    return file_path

def fetch_data_from_file(file_path):
    wb = open_workbook(file_path)
    for sheet in wb.sheets():
        values = []
        for row in range(sheet.nrows):
            col_value = []
            for col in range(sheet.ncols):
                value = (sheet.cell(row, col).value)
                try:
                    value = str(int(value))
                except:
                    pass
                col_value.append(value)
            values.append(col_value)
    return values

def load_session(data):
    table_columns = [Column('{}'.format(i), String) for i in data[0]]
    t = Table('t', metadata, *table_columns, Column('id', Integer, primary_key=True))
    metadata.create_all()
    conn = engine.connect()

    for value in data[1:]:
        ins = t.insert().values(value)
        conn.execute(ins)
    mapper(table_class, t)
    Session = sessionmaker(bind=engine)
    session = Session()
    return session

def run_queries(session):
    #to fetch all data
    res = session.query(table_class).all()
    for row in res:
        print(row.__dict__)

if __name__ == '__main__':
    file_path = get_file_path()
    data = fetch_data_from_file(file_path)
    session = load_session(data)
    run_queries(session)


