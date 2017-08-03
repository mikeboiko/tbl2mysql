# Name: xl2mysql.py
# Description: Convert Excel workbook into a MySQL database
# Authors: Mike Boiko, Travis Gall
# Notes:
# - Custom Excel file required and provided
#   - Contains a config worksheet with database credentials

# Imports{{{1

# Argument parsing
import argparse
# System functions
import os, sys
# MySQL connection
import pymysql
# Excel connection
import xlrd

# Global Variables{{{1

# Header row within Excel sheet
rowHeaders = 7

# Parser{{{1

# Create parser with script description
parser = argparse.ArgumentParser(description='Convert Excel workbook into a MySQL database')

# Add Excel workbook path
parser.add_argument(dest='xlWB', action='store', help='Full or relative path of workbook')

# Create objects from arguments passed and information within parser
args = parser.parse_args()

# Excel{{{1
# Open Connection{{{2

#  TODO [170803] - XML instead of Excel?
# Create connection to excel workbook without preloading sheet data
#  - This will improve initial load time
book = xlrd.open_workbook(args.xlWB, on_demand = True)

# Database Configuration{{{2

# Open worksheet with database configuration information
sheet = book.sheet_by_name('config')

# Extract database configuration
sqlHost = sheet.row_values(0)[1]
sqlPort = int(sheet.row_values(1)[1])
sqlUser = sheet.row_values(2)[1]
sqlPassword = sheet.row_values(3)[1]
sqlDB = sheet.row_values(4)[1]

# Queries{{{2

#  TODO [170803] - Restructure connection and query to handle using a new database
# Create database if required
sqlQuery = "CREATE DATABASE IF NOT EXISTS " + sqlDB + "; "

# Loop through all worksheets in workbook
for sheetName in book.sheet_names():

    #  TODO [170803] - Add options for update, insert and drop
    # Ignore config worksheet
    if sheetName != "config":
        sqlQuery += "DROP TABLE IF EXISTS " + sheetName + "; "
        sqlQuery += "CREATE TABLE " + sheetName + "(id INT NOT NULL AUTO_INCREMENT PRIMARY KEY"
        insert = "INSERT INTO " + sheetName + " ("

        # Load worksheet
        sheet = book.sheet_by_name(sheetName)

        #  TODO [170803] - Handlder required for column configurations (rows `range(1, 6)`)
        try:
            #  Loop through table headers
            for colNum in range(1, sheet.ncols):
                # Current cell value
                cellValue = sheet.cell(rowHeaders,colNum).value
                #  TODO [170803] - Column configuration
                # Add column name to CREATE query
                sqlQuery += ", " + cellValue + " text"
                #  TODO [170803] - Column configuration
                # Add column name to INSERT query
                insert += cellValue + ", "

            #  Begin first row of VALUES in query
            insert = insert[:-2] + ") VALUES ("
        except:
            # Print error to stdout
            print("Error in headers on sheet " + sheetName)
            raise
        finally:
            try:
                #  Loop through table rows
                for rowNum in range(8, sheet.nrows):
                    #  Begin new row of VALUES in query
                    if rowNum != 8: insert += ", ("

                    # Loop through all columns
                    for colNum in range(1, sheet.ncols):
                        # Current cell value
                        cellValue = sheet.cell(rowNum,colNum).value

                        # Cell has data: Add cell data to INSERT query
                        if cellValue: insert += "'" + str(int(cellValue)) + "', "
                        #  TODO [170803] - Column configuration
                        # NULL data: default data
                        else: insert += "'No Data', "
                    insert = insert[:-2] + ")"
            except:
                # Print error to stdout
                print("Error in data on sheet " + sheetName)

        # close insert values
        sqlQuery += "); " + insert + "; "

# Print MySQL query to stdout
print(sqlQuery)

# Close Connection{{{2

# Close workbook connection required for on_demand sheet data
book.release_resources()

# MySQL{{{1
db = pymysql.connect(host = sqlHost, port = sqlPort, user = sqlUser, passwd = sqlPassword, db = sqlDB)

try:
    # Exexute query
    db.cursor().execute(sqlQuery)

    # Commit all database modifications
    db.commit()
finally:
    db.close()

# Vim{{{1

# Custom fold method while using vim
# vim: foldmethod=marker:foldlevel=0
