# -*- coding: utf-8 -*-
"""
Created on Mon Feb  7 22:01:07 2022
Excel to Database converter
@author: icebox
"""
import csv
import sqlite3
import pandas as pd

def Excel2DBTables(ExcelFile, TableName):
    xldf = pd.read_csv(ExcelFile)
    xldf_header = xldf.columns.ravel()

    headers = []                        # Only generates TEXT columns
    for name in xldf_header:
        headers.append(name + " TEXT")
    headers = ", ".join(headers)
    
    cursor = connection.cursor()        # Create a cursor object to execute SQL queries on a database table
    # Table Definition
    create_table = 'CREATE TABLE ' + TableName + ' (id INTEGER PRIMARY KEY AUTOINCREMENT, ' + headers + ');'
    cursor.execute(create_table)        # Create the table into the database
    return cursor
    
def PopulateTables(ExcelFile, TableName, cursor):
    xldf = pd.read_csv(ExcelFile); xldf_header = xldf.columns.ravel()

    colnames = []; items = []   
    for name in xldf_header:                        # generate the header names
        colnames.append(name)
        items.append("?")
    colnames = ", ".join(colnames); items = ", ".join(items)

    # Populate the table
    with open(ExcelFile, encoding="utf-8") as csvfile:
        csvreader = csv.reader(csvfile)
        next(csvreader)                             # skips the first row of the CSV file
    
        # SQL query to insert data into the table
        insert_records = 'INSERT INTO ' + TableName + ' (' + colnames + ') ' + 'VALUES (' + items + ')'
     
        cursor.executemany(insert_records, csvreader) # Import the contents of the file into table
    return print ("\nPopulating database done!")

# =============================================================================
# Main
# =============================================================================
ExcelFile = 'UniqueCompDB_2022-03-29.csv'; # SheetName = 'UniqueCompDB_2022-03-29'
TableName = "company"; dbName = 'B2BVector_2022-03-30'

connection = sqlite3.connect(dbName + '.sqlite')    # Create and Connect to the database

cursor = Excel2DBTables(ExcelFile, TableName)
PopulateTables(ExcelFile, TableName, cursor)

connection.commit()                                 # Commit the changes
connection.close()                                  # close the database connection