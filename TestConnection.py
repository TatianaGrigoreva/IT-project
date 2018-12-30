# -*- coding: utf-8 -*-
"""
Created on Sun Dec 30 17:17:48 2018

@author: mikha
"""

import pyodbc # import pyodbc library for driver

server = 'DESKTOP-HP138CN\SQLEXPRESS' #server name
database = 'TestBase' #database name
user = 'PythonUser' #login
password = 'PythonUse' #password

connection = pyodbc.connect('Driver={SQL Server Native Client 11.0};SERVER=' + server 
                            + ';DATABASE=' + database
                            + ';UID=' + user 
                            + ';PWD=' + password
                            +' ;Trusted_Connection=yes') #create new connection with server
cursor = connection.cursor() #create special object for queries				  
cursor.execute('SELECT * FROM dbo.TestTable') #create new query
for row in cursor:
    print('row = %r' % (row,)) #print all rows in table
cursor.close()
connection.close() #close connection