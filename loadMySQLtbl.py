#!/usr/bin/python3
# =======================================================================
# === Description ...: Load table into MySQL database
# === Description ...: The table can be Excel or CSV based
# === Author ........: Mike Boiko and Travis Gall
# =======================================================================

# Imports {{{1

import argparse
import csv
import os
import pymysql
import sys
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

# Variables/Constants {{{1

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
    'Look at file extension to determine table type'

    # Declare global namespace variables
    global inputTableIsExcel
    global inputTableIsCSV

    # Initialize values
    inputTableIsExcel = False
    inputTableIsCSV = False

    # Check for the the input table file's extension
    # The . is added before extensiong and the string is made lower case
    # Excel Table
    if any('.'+substring in args.inputTableName.lower() for substring in fileExtensionsExcel):
        inputTableIsExcel = True

    # CSV Table
    elif any('.'+substring in args.inputTableName.lower() for substring in fileExtensionsCSV):
        inputTableIsCSV = True

    # Unknown file type
    else:
        sys.exit('Error: Unrecognizable File Extension for {}'.format(args.inputTableName))

def defineSqlDBTableName(): # {{{2
    'MySQL db Table Name'

    global sqlTableName

    # Defautlt table name - same as the input file name
    if args.sqlTableName == "default":
        # Strip out file name from path
        sqlTableName = os.path.basename(args.inputTableName)
        sqlTableName = os.path.splitext(sqlTableName)[0]
        # Only keep alpha-numeric characters
        sqlTableName = ''.join(e for e in sqlTableName if e.isalnum())

    # User defined table name in argument
    else:
        sqlTableName = args.sqlTableName

def initializeCSV(): # {{{2
    'Perform initialization for CSV file types'
    print('Parsing {} CSV table into {} MySql db'.format(args.inputTableName, args.database))

    global headerRow # Table field names

    with open(args.inputTableName, newline='') as fileCSV:
        reader = csv.reader(fileCSV)
        headerRow = next(reader)  # gets the first line

def initializeExcel(): # {{{2
    'Perform initialization for Excel file types'
    print('Parsing {} Excel table into {} MySql db'.format(args.inputTableName, args.database))

    global headerRow # Table field names
    global sheet     # Excel worksheet object

    excelWorkbookName = xlrd.open_workbook(args.inputTableName)
    if args.excelWorksheetName == 'default':
        sheet = excelWorkbookName.sheet_by_index(0)
    else:
        sheet = excelWorkbookName.sheet_by_name(args.excelWorksheetName)
    headerRow = sheet.row_values(0)

def prepareSqlQueries(): # {{{2
    'Prepare sql query strings'

    global sqlQueryCreate        # Create Table
    global sqlQueryInsert        # Insert into Table
    global sqlQueryInsertGeneric # Insert query generic string
    global sqlQueryTotal         # Combined queries

    # Prepare sql substrings that will be joined later
    sqlQueryInsert = ''
    sqlQueryDrop = f'drop table if exists {sqlTableName}; '

    sqlQueryCreate = f'create table {sqlTableName} (id int not null auto_increment primary key, '
    sqlInsertA = "insert into " + sqlTableName + " ("
    sqlInsertB = ""
    for header in headerRow:
        sqlQueryCreate += header + " text,\n"
        sqlInsertA += header + ", "
        sqlInsertB += r"'{}', "
    sqlQueryCreate = sqlQueryCreate[:-2] # Remove last ,
    sqlQueryCreate += '); '
    sqlInsertA = sqlInsertA[:-2] # Remove last ,
    sqlInsertB = sqlInsertB[:-2] # Remove last ,
    sqlQueryInsertGeneric = sqlInsertA + ") VALUES (" + sqlInsertB + ")"

    # TODO-MB [171125] - Re-write this script so the CSV/XL if statement doesn't have to happen twice

    # CSV - Prepare insert SQL query
    if inputTableIsCSV:
        sqlInsertDataFromCSV()
    # Excel - Prepare insert SQL query
    elif inputTableIsExcel:
        sqlInsertDataFromExcel()
    # There may be other types of tables other than Excel/CSV added later

    # All of the SQL queries combined into one string
    sqlQueryTotal = sqlQueryDrop + sqlQueryCreate + sqlQueryInsert

def sqlInsertDataFromCSV(): # {{{2
    'CSV - Prepare insert SQL query'

    global sqlQueryInsert # Insert into Table

    with open(args.inputTableName, newline='') as fileCSV:
        reader = csv.reader(fileCSV)
        next(reader) # Skip header line
        for row in reader:
            # Row needs to be converted from list to tuple and expanded with *
            sqlQueryInsert += sqlQueryInsertGeneric.format(*tuple(row)) + '; '

def sqlInsertDataFromExcel(): # {{{2
    'Excel - Prepare insert SQL query'

    global sqlQueryInsert # Insert into Table

    for rowNum in range(1, sheet.nrows):
        values = () # blank tuple
        for colNum in range(0, sheet.ncols):
            values = values + (sheet.cell(rowNum,colNum).value,)
        # Tuple needs to expanded with * for format function
        sqlQueryInsert += sqlQueryInsertGeneric.format(*tuple(values)) + '; '

def mySqlWrite(): # {{{2
    'Perform MySQL db write operations'

    # MySQL Connection
    db = pymysql.connect(host=args.host,
                         port=args.port,
                         user=args.user,
                         passwd=args.password,
                         db=args.database)

    try:
        # Exexute query
        db.cursor().execute(sqlQueryTotal)

        # Commit all database modifications
        db.commit()
    finally:
        db.close()

# Main Program {{{1

# Initialize File Extension lists for Excel and CSV
initFileExtensionLists()

# Look at file extension to determine table type
determineInputTableType()

# MySQL db Table Name
defineSqlDBTableName()

# Initialize CSV file types
if inputTableIsCSV:
    initializeCSV()
# Initialize Excel file types
elif inputTableIsExcel:
    initializeExcel()
# There may be other types of tables other than Excel/CSV added later

# Prepare sql query strings
prepareSqlQueries()

# Perform MySQL db write operations
mySqlWrite()

# os.system("pause")
# sys.exit()
