#!/usr/bin/python3
# =======================================================================
# === Description ...: Load table into MySQL database
# === Description ...: The tablse can be Excel or CSV based
# === Author ........: Mike Boiko and Travis Gall
# =======================================================================

# Imports {{{1

import argparse
import os, sys
import pymysql
import xlrd

# Parse Arguments {{{1

parser = argparse.ArgumentParser(description='Convert Excel Table to MySQL Table')

parser.add_argument(dest='inputTableName', action='store',
                    help='Name of input table file')
parser.add_argument('-ws', '--worksheet', dest='excelWorksheetName', action='store',
                    default='default',
                    help='Name of xlsx worksheet - default is 1st sheet')
parser.add_argument('--host', dest='host', action='store',
                    default='localhost',
                    help='MySQL server host address - default is localhost')
parser.add_argument('-p', '--port', dest='port', action='store',
                    default=3306,
                    help='MySQL port number')
parser.add_argument('-u', '--user', dest='user', action='store',
                    default='root',
                    help='MySQL username')
parser.add_argument('-pw', '--password', dest='password', action='store',
                    default='',
                    help='MySQL password')
parser.add_argument('-db', '--database', dest='database', action='store',
                    default='sample',
                    help='MySQL database name')
parser.add_argument('-t', '--table', dest='sqlTableName', action='store',
                    default='default',
                    help='MySQL table name - default is Workbook name')

args = parser.parse_args()

# Variables {{{1

# File extension string lists
fileExtensionsCSV = []
fileExtensionsExcel = []

# Functions {{{1
def initFileExtensionLists(): # {{{2
    "Initialize File Extension lists for Excel and CSV"

    # CSV possible file extensions
    fileExtensionsCSV.append("csv")
    fileExtensionsCSV.append("tsv")

    # Excel possible file extensions
    fileExtensionsExcel.append("xlsx")
    fileExtensionsExcel.append("xlsm")
    fileExtensionsExcel.append("xls")

def determineInputTableType(): # {{{2
    "Look at file extension to determine table type"

    # Declare global namespace variables
    global inputTableIsExcel
    global inputTableIsCSV

    # Check for the the input table file's extension
    # The . is added before extensiong and the string is made lower case
    if any('.'+substring in args.inputTableName.lower() for substring in fileExtensionsExcel):
        inputTableIsExcel = True
        inputTableIsCSV = False
    elif any('.'+substring in args.inputTableName.lower() for substring in fileExtensionsCSV):
        inputTableIsExcel = False
        inputTableIsCSV = True
    else:
        sys.exit('Error: Unrecognizable File Extension for {}'.format(args.inputTableName))


# Main Program {{{1

# Initialize File Extension lists for Excel and CSV
initFileExtensionLists()

# Look at file extension to determine table type
determineInputTableType()

print(args.inputTableName)
# print('Excel Flag is {}, CSV flag is {}'.format(inputTableIsExcel, inputTableIsCSV))
# os.system("pause")

if inputTableIsCSV:
    print("CSV")
else:
    print("Excel")

os.system("pause")
sys.exit()

# Excel Connection {{{2
excelWorkbookName = xlrd.open_workbook(args.inputTableName)
if args.excelWorksheetName == 'default':
    sheet = excelWorkbookName.sheet_by_index(0)
else:
    sheet = excelWorkbookName.sheet_by_name(args.excelWorksheetName)

# MySQL Connection {{{2
db = pymysql.connect(host=args.host,
                     port=args.port,
                     user=args.user,
                     passwd=args.password,
                     db=args.database)
cursor = db.cursor()

# MySQL Table Name {{{2
if args.sqlTableName == "default":
    sqlTableName = args.inputTableName.replace('.xlsx','').replace('.xlsm','').replace('.xls','')
    # Strip special characters from Excel file name
    sqlTableName = ''.join(e for e in sqlTableName if e.isalnum())
else:
    sqlTableName = args.sqlTableName

# Prepare Queries. {{{2
# sql1: create table, sql2: insert
excelHeaders = sheet.row_values(0)
sql1 = "create table " + sqlTableName + " (id int not null auto_increment primary key, "
sql2a = "insert into " + sqlTableName + " ("
sql2b = ""
for header in excelHeaders:
    sql1 += header + " text,\n"
    sql2a += header + ", "
    sql2b += "%s, "
sql1 = sql1[:-2] # Remove last ,
sql1 += ")"
sql2a = sql2a[:-2] # Remove last ,
sql2b = sql2b[:-2] # Remove last ,
sql2 = sql2a + ") VALUES (" + sql2b + ")"

# Modify Table {{{2

# Create Table
cursor.execute("drop table if exists " + sqlTableName)
cursor.execute(sql1)

# Insert all data into table
for rownum in range(1, sheet.nrows):
    values = () # blank tuple
    for colnum in range(0, sheet.ncols):
        values = values + (sheet.cell(rownum,colnum).value,)
    cursor.execute(sql2, values)

# Close out {{{2
cursor.close()
db.commit()
db.close()
