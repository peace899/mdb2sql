# Author: Peace Lekalakala
# peacester at gmail dot com

import io
import os
import pandas as pd
import pyodbc
import sys

from flask import Flask, request, send_file
from os.path import expanduser
from werkzeug.utils import secure_filename

home = expanduser("~")
UPLOAD_FOLDER = os.path.join(home, 'FlaskData')

try:
    os.mkdir(UPLOAD_FOLDER)
except OSError:
    pass
    
ALLOWED_EXTENSIONS = set(['MDB', 'mdb'])

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def find_drv(name, path):
    for root, dirs, files in os.walk(path):
        if name in files:
            return os.path.join(root, name)

def convert_to_xls(mdb_file, xlsf):
    writer = pd.ExcelWriter(xlsf, engine='xlsxwriter')   
    
    # Connect to mdb
    if sys.platform.startswith('linux'):
        if os.path.exists('/usr/lib/libmdbodbc.so'):
            DRV = '/usr/lib/libmdbodbc.so'
        else:
            DRV = str(find_drv('libmdbodbc.so', '/'));
        MDB = mdb_file.replace(" ", "\ ")
        conn = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV, MDB))
        conn.autocommit = True
        conn.setencoding('latin1')
        conn.setdecoding(pyodbc.SQL_CHAR, 'latin1')
        conn.setdecoding(pyodbc.SQL_WCHAR, 'latin1')
        cur = conn.cursor()
    else:
        DRV = '{Microsoft Access Driver (*.mdb)}';
        MDB = mdb_file
        conn = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV,MDB))
        cur = conn.cursor()
          
    # Get table names from the mdb
    tables = []
    for row in cur.tables():
        if 'MSys' not in row.table_name:
            tables.append(row.table_name)
    
    # Create worksheets from tables
    for table in tables:
        print('Creating worksheet {}'.format(table))
        
        # Query table
        if sys.platform.startswith('linux'):
            SQL = 'SELECT * FROM {}'.format(table) 
        else:
            SQL = 'SELECT * FROM [{}];'.format(table) 
        
        df = pd.read_sql(SQL, conn)
        df.to_excel(writer, sheet_name=table, index=False, encoding="utf-8")
        
    # Save 'final' excel file and close mdb connection    
    writer.save()
    cur.close()
    return xlsf
    
    
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            
            # Save the file to convert in upload folder
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            mdb_file = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            
            # Name final excel file with mdb prefix
            db_file_short = os.path.basename(mdb_file)
            xls_file = os.path.splitext(db_file_short)[0]+ '.xls'
            
            # Create an excel file in memory
            mem_file = io.BytesIO()
            
            # Convert mdb file to excel format and read from memory
            xlsf = convert_to_xls(mdb_file, mem_file)
            xlsf.seek(0)
            
            # Remove the uploaded file
            os.remove(mdb_file)
            
            # Send excel file as download after converting
            return send_file(xlsf,
                             attachment_filename=xls_file,
                             as_attachment=True)
            
        else:
            return "Please choose MSAccess mdb file"
    return '''
    <!doctype html>
    <title>MDB2XLS Converter</title>
    <h1>Convert mdb file to xls</h1>
    <form action="" method=post enctype=multipart/form-data>
      <p><input type=file name=file></p>
      <p><input type=submit value=Convert></p>
    </form>
    '''
if __name__ == "__main__":
    app.run(host='127.0.0.1', port=5000, debug=True)
