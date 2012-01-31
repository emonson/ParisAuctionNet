#!/usr/bin/env python

"""
Script to move articles (e.g. de, de la, d'...) from before
last name to after first name in Hilary's FileMaker Pro database
through the ODBC connection

Takes in an Excel workbook with three sheets:
	one for the list of articles
	one for the column and table names (Table, ID name, Last name, First name)
	one for name exceptions that should not be changed
	
31 Jan 2012 -- E Monson
"""

import pyodbc
from openpyxl import load_workbook

def xflatten(seq):
	"""a generator to flatten a nested list
	from vegaseat on http://www.daniweb.com/forums/thread66694.html
	usage: flat_list = list(xflatten(nested_list))
	"""
	for x in seq:
		if type(x) in [list, tuple]:
			for y in xflatten(x):
				yield y
		else:
				yield x

def n_str(s, a):
	"""Deals with None in first_name"""
	if s is None:
		return unicode(a.strip())
	else:
		return unicode(s.decode('utf8').strip() + ' ' + a.strip())

# This ODBC connection has to be set up ahead of time
cnxn = pyodbc.connect('DSN=ParisAuctionNet;UID=emonson;PWD=test.log')
cursor = cnxn.cursor()

in_file = 'article_details.xlsx'

wb = load_workbook(in_file)

multi_columns_sheet = wb.get_sheet_by_name("multi_columns")
single_columns_sheet = wb.get_sheet_by_name("single_columns")
articles_sheet = wb.get_sheet_by_name("articles")
except_sheet = wb.get_sheet_by_name("exceptions")

multi_columns = [tuple(xx.value for xx in yy) for yy in multi_columns_sheet.rows]
single_columns = [tuple(xx.value for xx in yy) for yy in single_columns_sheet.rows]
articles = [xx.value.replace('_',' ') for xx in list(xflatten(articles_sheet.rows))]
exceptions = [xx.value for xx in list(xflatten(except_sheet.rows))]

for ns in multi_columns:
	# ns = (id, last_name, first_name, table)
	print 'TABLE/NAME:', ns
	
	for aa in articles:
		print 'ARTICLE:', aa
	
		# Having trouble with pyodbc parameter substitution, and no external user input,
		# so not going to worry about injection...
		params = (ns[0], ns[1], ns[2], ns[3], ns[1])
		query = (u'SELECT "%s" as id, "%s" as last_name, "%s" as first_name FROM "%s" WHERE "%s" LIKE ?' % params)
		
		# Need to be careful with the order of the articles
		# since some are subsets of others. Longest first of any similar items
		# (e.g. do "de la " before "de " or will end up with "la " last name beginnings
		#  not matched with "de la ")
		c1 = cursor.execute(query, aa + u'%').fetchall()
		# print c1
		
		if len(c1) > 0:

			c2 = [(unicode(last_name.decode('utf8').split(aa)[-1].strip()), n_str(first_name, aa), id) for (id, last_name, first_name) in c1 if last_name not in exceptions]
			# print c2
		
			params2 = (ns[3], ns[1], ns[2], ns[0])
			update_str = (u'UPDATE "%s" SET "%s" = ?, "%s" = ? WHERE "%s" = ?' % params2)
			
			c3 = cursor.executemany(update_str, c2)
			print 'UPDATED:', len(c2)
			cnxn.commit()

for ns in single_columns:
	# ns = (id, name, table)
	print 'TABLE/NAME:', ns
	
	for aa in articles:
		print 'ARTICLE:', aa
	
		# Having trouble with pyodbc parameter substitution, and no external user input,
		# so not going to worry about injection...
		params = (ns[0], ns[1], ns[2], ns[1])
		query = (u'SELECT "%s" as id, "%s" as name FROM "%s" WHERE "%s" LIKE ?' % params)
		
		# Need to be careful with the order of the articles
		# since some are subsets of others. Longest first of any similar items
		# (e.g. do "de la " before "de " or will end up with "la " last name beginnings
		#  not matched with "de la ")
		c1 = cursor.execute(query, aa + u'%').fetchall()
		# print c1
		
		if len(c1) > 0:

			c2 = [(unicode(name.decode('utf8').split(aa)[-1].strip() + ' ' + aa.strip()), id) for (id, name) in c1 if name not in exceptions]
			# print c2
		
			params2 = (ns[2], ns[1], ns[0])
			update_str = (u'UPDATE "%s" SET "%s" = ? WHERE "%s" = ?' % params2)
			
			c3 = cursor.executemany(update_str, c2)
			print 'UPDATED:', len(c2)
			cnxn.commit()


cnxn.close()
