#!/usr/bin/python3
# =======================================================================
# === Description ...: Load Excel table into MySQL database
# === Author ........: Mike Boiko and Travis Gall
# =======================================================================

import argparse
import os, sys
import pymysql
import xlrd

# =============================================
# === Parse Arguments
# =============================================

parser = argparse.ArgumentParser(description='Convert Excel Table to MySQL Table')

parser.add_argument(dest='xlWB', action='store',
                    help='Name of xlsx workbook')
parser.add_argument('-ws', '--worksheet', dest='xlWS', action='store',
                    default='default',
                    help='Name of xlsx worksheet - default is 1st sheet')
parser.add_argument('-i', '--ip', dest='ip', action='store',
                    default='localhost',
                    help='MySQL server ip address')
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
parser.add_argument('-t', '--table', dest='table', action='store',
                    default='default',
                    help='MySQL table name - default is Workbook name')

args = parser.parse_args()

# =============================================
# === Main
# =============================================

# Establish Excel connection
book = xlrd.open_workbook(args.xlWB)
if args.xlWS == 'default':
    sheet = book.sheet_by_index(0)
else:
    sheet = book.sheet_by_name(args.xlWS)

# Establish MySQL connection
db = pymysql.connect(host=args.ip,
                port=args.port,
                user=args.user,
                passwd=args.password,
                db=args.database)
cursor = db.cursor()

# MySQL Table Name
if args.table == "default":
    tableName = args.xlWB.replace('.xlsx','').replace('.xlsm','').replace('.xls','')
    # Strip special characters from Excel file name
    tableName = ''.join(e for e in tableName if e.isalnum())
else:
    tableName = args.table

# Prepare Queries. sql1: create table, sql2: insert
excelHeaders = sheet.row_values(0)
sql1 = "create table " + tableName + " (id int not null auto_increment primary key, "
sql2a = "insert into " + tableName + " ("
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

# Create table
cursor.execute("drop table if exists " + tableName)
cursor.execute(sql1)

# Insert all data into table
for rownum in range(1, sheet.nrows):
    values = () # blank tuple
    for colnum in range(0, sheet.ncols):
        values = values + (sheet.cell(rownum,colnum).value,)
    cursor.execute(sql2, values)

# Close out
cursor.close()
db.commit()
db.close()
