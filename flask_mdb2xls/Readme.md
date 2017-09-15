This script converts TIPS MS Access .mdb writeout files to Excel formats that can be imported as DGS format into DigSilent Powerfactory.

This script requires the following modules and libraries installed:
 * [PyODBC](https://mkleehammer.github.io/pyodbc/) 
 * [xlwt](https://pypi.python.org/pypi/xlwt)
 * [Flask web framework](http://flask.pocoo.org/)
 * [Werkzeug : Web Server Gateway Interface (WSGI)](http://werkzeug.pocoo.org/)

To install the libraries and modules just run:
  > pip install pyodbc
  
  > pip install xlwt
  
  > pip install flask
  
  > pip install Werkzeug



 
To use just run the python script and open your browser and go to address: http://127.0.0.1:5000
Browse for the file you want convert and it sends you an Excel file as download.

 
Preview:
![](mdb2xls.PNG?raw=true)
 
