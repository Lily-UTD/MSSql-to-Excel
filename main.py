# pip install pyodbc
# pip install openpyxl
from SQLHandler import sqlhandler
from ExcelHandler import excelhandler
from SQLFileHandler import sqlfilehandler

sqlfile = sqlfilehandler('query.sql')
sheetnames = sqlfile.getSheetnames()
queries = sqlfile.getQueries()
sql = sqlhandler('192.168.1.1','usr','pwd','dbname')
for sqlcmd, sheetname in zip(queries, sheetnames):
    sql.exeCommand(sqlcmd)
    header = sql.getHeader()
    data = sql.getData()
    excel = excelhandler('test.xlsx',sheetname)
    excel.setHeader(header)
    excel.setData(data)
    excel.saveExcl()
