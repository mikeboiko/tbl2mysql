# =======================================================================
# === Description ...: Load table into MySQL database
# ===             ...: The table can be Excel or CSV based
# === Authors .......: Mike Boiko and Travis Gall
# =======================================================================

# Imports {{{1

import argparse # Argument parsing
import csv      # CSV read/write
import os       # OS interface
import pymysql  # MySQL connection
import sys      # System functions
import xlrd     # Excel Connection

# Parse Arguments {{{1

# Create parser with script description
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
                    help='MySQL table name - default is the input name')
parser.add_argument('--dropTable', action='store_true',
                    help='Drop the old table if it exists.')

# Create objects from arguments passed and information within parser
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
        # Only keep alpha-numeric characters and underscores
        sqlTableName = ''.join(e for e in sqlTableName if e.isalnum() or '_' in e)

    # User defined table name in argument
    else:
        sqlTableName = args.sqlTableName

def initializeCSV(): # {{{2
    'Perform initialization for CSV file types'

    global headerRow # Table field names

    # Open CSV and read first line
    with open(args.inputTableName, newline='') as fileCSV:
        reader = csv.reader(fileCSV)

        # Get header row (used for extracting field names)
        try:
            headerRow = next(reader)
        except:
            sys.exit(f'Error: {args.inputTableName} is empty')
            # raise

def initializeExcel(): # {{{2
    'Perform initialization for Excel file types'

    global headerRow     # Table field names
    global sheet         # Excel worksheet object
    global excelWorkbook # Excel book object

    # Create connection to excel workbook without preloading sheet data
    #  - This will improve initial load time
    excelWorkbook = xlrd.open_workbook(args.inputTableName, on_demand = True)

    # User either specifies a worksheet by name or the first sheet is used
    if args.excelWorksheetName == 'default':
        sheet = excelWorkbook.sheet_by_index(0)
    else:
        sheet = excelWorkbook.sheet_by_name(args.excelWorksheetName)

    # Get header row (used for extracting field names)
    headerRow = sheet.row_values(0)

def sqlQueriesPrepare(): # {{{2
    'Prepare sql query strings'

    global sqlQueryCreate        # Create table query
    global sqlQueryDrop          # Drop table query
    global sqlQueryInsert        # Insert into Table query
    global sqlQueryInsertGeneric # Insert query generic string

    # Prepare sql substrings that will be joined later
    sqlQueryInsert = ''
    sqlQueryDrop = f'DROP TABLE IF EXISTS {sqlTableName}; '

    sqlQueryCreate = f'''CREATE TABLE IF NOT EXISTS {sqlTableName}
                     (id INT NOT NULL AUTO_INCREMENT PRIMARY KEY, '''
    sqlInsertA = f'INSERT INTO {sqlTableName} ('
    sqlInsertB = ''
    for header in headerRow:
        sqlQueryCreate += f'{header} TEXT, '
        sqlInsertA += f'{header} , '
        sqlInsertB += r"'{}', "
    sqlQueryCreate = sqlQueryCreate[:-2] # Remove last ,
    sqlQueryCreate += '); '
    sqlInsertA = sqlInsertA[:-2] # Remove last ,
    sqlInsertB = sqlInsertB[:-2] # Remove last ,
    sqlQueryInsertGeneric = f'{sqlInsertA}) VALUES ({sqlInsertB})'

    # TODO-MB [171125] - Re-write this script so the CSV/XL if statement only occurs one time

    # CSV - Prepare insert SQL query
    if inputTableIsCSV:
        sqlInsertDataFromCSV()
    # Excel - Prepare insert SQL query
    elif inputTableIsExcel:
        sqlInsertDataFromExcel()
    # There may be other types of tables other than Excel/CSV added later

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

    # Loop through each cell
    for rowNum in range(1, sheet.nrows):
        values = () # Initialize blank tuple

        for colNum in range(0, sheet.ncols):
            values = values + (sheet.cell(rowNum,colNum).value,)

        # Tuple needs to expanded with * for format function
        sqlQueryInsert += sqlQueryInsertGeneric.format(*tuple(values)) + '; '

    # Close workbook connection required for on_demand sheet data
    excelWorkbook.release_resources()

def mySqlDbConnect(): # {{{2
    'Initialize db connection'

    global db # Database object

    # MySQL Connection
    db = pymysql.connect(host=args.host,
                         port=args.port,
                         user=args.user,
                         passwd=args.password,
                         db=args.database)

    # Supress annoying warnings
    from warnings import filterwarnings
    filterwarnings('ignore', category = db.Warning)

def sqlDropCreateTable(): # {{{2
    'Drop old table/Create new table if required'

    global sqlQueryCreate # Create table query
    global sqlQueryDrop   # Drop table query

    # Only drop table if the user requested it
    if not args.dropTable:
        sqlQueryDrop = ''

    # Combine SQL Queries
    sqlQueryTotal = sqlQueryDrop + sqlQueryCreate

    try:
        # Exexute query
        db.cursor().execute(sqlQueryTotal)
    except:
        # Print error to stdout
        print("Unexpected error:", sys.exc_info()[0])
        raise

def sqlAddNewFields(tableName): # {{{2
    'When appending to existing table, add new fields if required'

    # Create cursor
    dbCur = db.cursor()

    # Query the existing table for 1 row of data only
    dbCur.execute(f'''
        SELECT *
        FROM {args.database}.{tableName}
        LIMIT 1
        ''')

    # Find if there are any fields in the input table that don't exist in the db table
    numFields = len(dbCur.description)
    fieldNames = [i[0] for i in dbCur.description]
    newFields = list(set(headerRow) - set(fieldNames))

    # If required, add new fields to db table
    if len(newFields) > 0:
        sqlQueryAlter = f'ALTER TABLE {sqlTableName} '
        for field in newFields:
            sqlQueryAlter += f'ADD COLUMN {field} TEXT, '
        sqlQueryAlter = sqlQueryAlter[:-2] # Remove last ,

        # Run alter query
        try:
            dbCur.execute(sqlQueryAlter)
        except:
            # Print error to stdout
            print("Unexpected error:", sys.exc_info()[0])
            raise

    # Close cursor
    dbCur.close()

def sqlInsertRecords(): # {{{2
    'Perform MySQL db write operations'

    try:
        # Exexute query
        db.cursor().execute(sqlQueryInsert)
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

# Initialize db connection
mySqlDbConnect()

# Prepare sql query strings
sqlQueriesPrepare()

# Drop old table/Create new table if required
sqlDropCreateTable()

# If appending to existing table, add new fields (if required)
if not args.dropTable:
    sqlAddNewFields(args.sqlTableName)

# Perform MySQL db write operations
sqlInsertRecords()

# Exit with no errors
sys.exit(0)

# os.system("pause")
# sys.exit()
