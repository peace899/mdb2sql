import os
import pyodbc
import sys
import time
import xlwt

from flask import Flask, request, redirect, url_for
from flask import send_from_directory, send_file
from werkzeug.utils import secure_filename

from os.path import expanduser
home = expanduser("~")
UPLOAD_FOLDER = os.path.join(home, 'FlaskData')

try:
    os.mkdir(UPLOAD_FOLDER)
except OSError:
    pass
    
ALLOWED_EXTENSIONS = set(['MDB', 'mdb'])

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def convert_to_xls(mdb_file, xls_file):
         
    DRV = '{Microsoft Access Driver (*.mdb)}'; 
    conn = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV,mdb_file))
    cur = conn.cursor()
    wb = xlwt.Workbook(encoding='utf-8')
    tables = []
    for row in cur.tables():
        if 'MSys' not in row.table_name:
            tables.append(row.table_name)
    
    for table in tables:
        print('Creating worksheet {}'.format(table))
        
                   
        SQL = 'SELECT * FROM [{}];'.format(table) # your query goes here
        rows = cur.execute(SQL).fetchall()
        columns = [i[0] for i in cur.description]
       
        
        ws = wb.add_sheet(table)
        row_num = 0
        font_style = xlwt.XFStyle()
        font_style.font.bold = True
        for col_num in range(len(columns)):
            ws.write(row_num, col_num, columns[col_num], font_style)
        
        for row in rows:
            row_num += 1
            for col_num in range(len(row)):
                ws.write(row_num, col_num, row[col_num])
        
    wb.save(xls_file)
    cur.close()
    
    
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            
            mdb_file = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            db_file_short = os.path.basename(mdb_file)
            xls_file = os.path.splitext(db_file_short)[0]+ '.xls'
            xls_file = os.path.join(app.config['UPLOAD_FOLDER'], xls_file)
            
            # Convert mdb file to excel
            convert_to_xls(mdb_file, xls_file)
            
            # Send excel file as download after converting
            return send_file(xls_file, as_attachment=True)
        else:
            return "file not allowed"
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
    app.run(host='0.0.0.0', port=5000, debug=True)
