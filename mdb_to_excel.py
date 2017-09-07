from __future__ import print_function
import pandas as pd
import numpy as np
import os
import json
from argparse import ArgumentParser
from gooey import Gooey, GooeyParser
import pyodbc
import time



@Gooey(program_name="Create Excel format DGS file from TIPS MDB writeout")
def parse_args():
    """ Use GooeyParser to build up the arguments we will use in our script
    Save the arguments in a default json file so that we can retrieve them
    every time we run the script.
    """
    stored_args = {}
    # get the script name without the extension & use it to build up
    # the json filename
    script_name = os.path.splitext(os.path.basename(__file__))[0]
    args_file = "{}-args.json".format(script_name)
    # Read in the prior arguments as a dictionary
    if os.path.isfile(args_file):
        with open(args_file) as data_file:
            stored_args = json.load(data_file)
    parser = GooeyParser(description='Create Excel workbook from MDB')
    parser.add_argument('mdb_file',
                        action='store',
                        default=stored_args.get('mdb_file'),
                        widget='FileChooser',
                        help="Source .mdb file to convert")
    parser.add_argument('output_directory',
                        action='store',
                        widget='DirChooser',
                        default=stored_args.get('output_directory'),
                        help="Output directory to save final Excel file")
    
    args = parser.parse_args()
    # Store the values of the arguments so we have them next time we run
    with open(args_file, 'w') as data_file:
        # Using vars(args) returns the data as a dictionary
        json.dump(vars(args), data_file)
    return args

            
def convert_to_xls(mdb_file, xls_file):
    writer = pd.ExcelWriter(xls_file)     
    DRV = '{Microsoft Access Driver (*.mdb)}'; 
    conn = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV,mdb_file))
    cur = conn.cursor()

    tables = []
    for row in cur.tables():
        if 'MSys' not in row.table_name:
            tables.append(row.table_name)
    
    for table in tables:
        print('Processing {} '.format(table))
                        
        SQL = 'SELECT * FROM [{}];'.format(table) # your query goes here
        df = pd.read_sql(SQL, conn)
        df.to_excel(writer,sheet_name=table,index=False)
    writer.save()
    cur.close()
    

if __name__ == '__main__':
    conf = parse_args()
    db_file_short = os.path.basename(conf.mdb_file)
    xls_file = os.path.splitext(db_file_short)[0]+ '.xls'
    xls_file = os.path.join(conf.output_directory, xls_file)
    print("Connecting to mdb file...{}".format(db_file_short))
    time.sleep(1.5)
    print('Connected to mdb file')
    print('Creating excel file...')
    convert_to_xls(conf.mdb_file, xls_file)
    print('Your Excel DGS file saved as: {}'.format(xls_file))
    
 
