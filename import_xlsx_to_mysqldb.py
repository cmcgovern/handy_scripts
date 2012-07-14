#!/usr/bin/env python
# encoding: utf-8
"""
import_xlsx_to_mysqldb.py

Assumes database has been created for the import and that there are never more headers than values in columns.

Will strip out spaces from headers to make valid field

Created by Colin McGovern on 2012-07-14.
Copyright (c) 2012 - Free to use, no warrantee offered, etc...

"""

import sys
import getopt
import MySQLdb
from xlsx import Workbook
from excel_handy_functions import convert_column_to_integer

myAppName = sys.argv[0].split("/")[-1]
help_message = "Usage: " + myAppName + " -f <excel file> -n <database name> -d <database server> -s <socket> -u <username> -p <password> [-x (Undo)]"

class Usage(Exception):
	def __init__(self, msg):
		self.msg = msg

# For the moment, just remove the spaces
def sql_compatible_name(freeform_name):
	return freeform_name.replace(" ", "_")

def utf_fix(string):
	"""tmp solution for latin-1 problem"""
	return string.encode('latin-1', 'replace')

def headers_for_worksheet(worksheet):
	return [sql_compatible_name(column[1][0].value) for column in sorted(worksheet.cols().items())]

# Columns above sorted to solve problem that the wrong header returning here
# Seems to be an issue with xlsx that it doesn't guarantee column order...
def available_headers_for_cells(cells, all_column_headers):
	return [sql_compatible_name(all_column_headers[convert_column_to_integer(cell.column)]) for cell in cells]

def execute_single_statement(database, statement, values = None):
	cursor = database.cursor()
	if values:
		cursor.execute(statement, values)
	else:
		cursor.execute(statement)
	cursor.close()
	database.commit()

def create_table(database, table, fields):
	print ("Creating table: %s, with fields: %s") % (table, fields)
	create_statement = ("CREATE TABLE IF NOT EXISTS %s (`%s` VARCHAR(256))") % (table, "` VARCHAR(256), `".join(fields))
	execute_single_statement(database, create_statement)

def drop_table(database, table):
	drop_statement = ("DROP TABLE IF EXISTS %s") % (table)
	execute_single_statement(database, drop_statement)
	
def insert_values(database, table, fields, values):
	value_placeholders = ["%s" for value in values]
	insert_statement = "INSERT INTO " + table + " (`" + "`,`".join(fields) + "`) VALUES (" + ",".join(value_placeholders) + ")"
	print ("Inserting %s using %s") % (values, insert_statement)
	execute_single_statement(database, insert_statement, values)

def create_table_with_worksheet(database, worksheet):
	"""Creates table using first row as headers"""
	
	# Create Table
	table_name = sql_compatible_name(worksheet.name)
	column_headers = headers_for_worksheet(worksheet)
	create_table(database, table_name, column_headers)
	
	# Populate Table
	for row, cells in worksheet.rows().items():
		if cells[0].row != 1: # Skip headers
			# Order can be wrong, so you must use the header to find the right value
			values = [utf_fix(cell.value) for cell in cells]
			insert_values(database, table_name, available_headers_for_cells(cells, column_headers), values)

def remove_table_for_worksheet(database, worksheet):
	"""Drops table using first row as headers"""
	table_name = sql_compatible_name(worksheet.name)
	drop_table(database, table_name)

def start_with_options(excel_file, name, server, socket, username, password, undo):
	verb = "Removing" if undo else "Importing"
	print >> sys.stdout, verb + " " + excel_file + " into " + name + "@" + server
	
	# connect to database
	try:
		if socket:
			db = MySQLdb.connect(host=server, unix_socket=socket, user=username, passwd=password, db=name)
		else:
			db = MySQLdb.connect(host=server, db=name, user=username, passwd=password)
			
		# Create a table for each worksheet
		for worksheet in Workbook(excel_file):
			if (undo):
				remove_table_for_worksheet(db, worksheet)
			else:		
				create_table_with_worksheet(db, worksheet)
			
	except MySQLdb.MySQLError, err:
		print >> sys.stderr, myAppName + ": " + str(err)
		return -1
	
		
def main(argv=None):
	if argv is None:
		argv = sys.argv
	try:
		try:
			opts, args = getopt.getopt(argv[1:], "hf:n:d:s:u:p:vx", ["help", "file=", "dbname=", "dbserver=", "dbsocket=", "dbuser=", "dbpassword=", "verbose", "undo"])
		except getopt.error, msg:
			raise Usage(msg)
	
		# Scope...
		excel_file = None
		name = None
		server = None
		socket = None
		username = None
		password = None
		undo = False
	
		# option processing
		for option, value in opts:
			if option in ("-v", "--verbose"):
				verbose = True
			if option in ("-h", "--help"):
				raise Usage(help_message)
			if option in ("-f", "--file"):
				excel_file = value
			if option in ("-n", "--dbname"):
				name = value
			if option in ("-d", "--dbserver"):
				server = value
			if option in ("-s", "--dbsocket"):
				socket = value
			if option in ("-u", "--dbuser"):
				username = value
			if option in ("-p", "--dbpassword"):
				password = value
			if option in ("-x", "--undo"):
				undo = True
				
		# Make sure we're not missing anything important
		if excel_file == None:
			raise Usage("Need excel file...")
			
		if name == None:
			raise Usage("Need database name")
			
		if server == None:
			raise Usage("Need database server...")
		
		if username == None or password == None:
			raise Usage("Need credentials...")
	
		start_with_options(excel_file, name, server, socket, username, password, undo)
	
	except Usage, err:
		print >> sys.stderr, myAppName + ": " + str(err.msg)
		print >> sys.stderr, "\t for help use --help"
		return 2


if __name__ == "__main__":
	sys.exit(main())
