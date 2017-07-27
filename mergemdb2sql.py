import os, sys
import glob
import csv, pyodbc
import pandas
import sqlite3

currdir = os.getcwd()
csv_dir = os.path.join(currdir, 'csv')
mdbs_path = r'path/to/mdbs'
db_file = os.path.join(currdir, 'test.db')


def create_csv(mdb_file):
    MDB = mdb_file
    if sys.platform.startswith('linux'):
        if os.path.exists('/usr/lib/libmdbodbc.so'):
            DRV = '/usr/lib/libmdbodbc.so'
        else:
            DRV = str(find_drv('libmdbodbc.so', '/'));
    else:
        DRV = '{Microsoft Access Driver (*.mdb)}';
        
    conn = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV,MDB))
    """
    conn.autocommit = True
    conn.setencoding('latin1')
    conn.setdecoding(pyodbc.SQL_CHAR, 'latin1')
    conn.setdecoding(pyodbc.SQL_WCHAR, 'latin1')
    """
    cursor = conn.cursor()
    
    tables = []
    for row in cursor.tables():
        if 'MSys' not in row.table_name:
            tables.append(row.table_name)
    
    for table in tables:    
        csv_file = os.path.join(csv_dir, '{}.csv'.format(table))
        SQL = 'SELECT * FROM [{}];'.format(table) # your query goes here
        rows = cursor.execute(SQL).fetchall()
        field_names = [i[0] for i in cursor.description]
        if not os.path.isfile(csv_file):
            with open(csv_file, 'wb') as fou:
                csv_writer = csv.writer(fou)
                csv_writer.writerow(field_names)
                csv_writer.writerows(rows)
            fou.close()
            
        else:
            with open(csv_file, 'a') as fou:
                csv_writer = csv.writer(fou) # default field-delimiter is ","
                #csv_writer.writerow(field_names)
                csv_writer.writerows(rows)
            fou.close()
        
    cursor.close()
    

def csv_to_sql(path):
    os.chdir(path)
    cnx = sqlite3.connect(db_file)
    for filename in glob.glob("*.csv"):
        (f_path, f_name) = os.path.split(filename)
        (f_short_name, f_extension) = os.path.splitext(f_name)
        tbl = str(f_short_name)
        df = pandas.read_csv(filename)
        df.to_sql(tbl, cnx)
       

def del_csv(path):
    os.chdir(path)
    for filename in glob.glob("*.csv"):
        os.remove(filename)
        

mdbs = []                                                              
for f in os.listdir(mdbs_path):
    if f.endswith(".mdb"):
        mdbs.append(os.path.join(mdbs_path, f))

for mdb in mdbs:
    create_csv(mdb)
         
csv_to_sql(csv_dir)
del_csv(csv_dir)
