# pip install pyodbc

import pyodbc

class sqlhandler(object):
    def __init__(self, dbip, dbusr, dbpwd, dbname):
        self.cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=' + dbip + ';DATABASE=' + dbname + ';UID=' + dbusr + ';PWD='+dbpwd)
        self.cursor = self.cnxn.cursor()
    
    def exeCommand(self, command):
        self.cursor.execute(command)
    
    def getHeader(self):
        self.rows = self.cursor.fetchall()
        columnheader = []
        for column in self.rows[0].cursor_description:
            columnheader.append(column[0])
        return columnheader
    
    def getData(self):
        data = []
        for row in self.rows:
            data.append(row)
        return data