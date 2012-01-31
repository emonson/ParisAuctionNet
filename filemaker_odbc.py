import pyodbc
cnxn = pyodbc.connect('DSN=ParisAuctionNet;UID=emonson;PWD=test.log')
cursor = cnxn.cursor()
# cursor.execute('select "Year of Sale","Date of Sale" from Catalogues')
# rows = cursor.fetchall()
# for row in rows:
#     print row[0]
# 
# cursor.execute('select Subject from Transactions')
# rows = cursor.fetchall()
# # looks like the original is in 'latin-1', but can use
# # this to shift to utf-8 unicode strings
# subjects = [x[0].decode('utf-8') for x in rows]

