import pyodbc
import nltk
import codecs
import re
from collections import defaultdict
import os

out_dir = '/Users/emonson/Data/ArtMarkets/Hilary/LDA'

# Regex used later to filter out "bad" words
ncends_re = re.compile(r'^[^a-zA-Z0-9]*([a-zA-Z0-9]+)[^a-zA-Z0-9]*$')
ncall_re = re.compile(r'^[^a-zA-Z]+$')
ncany_re = re.compile(r'[^a-zA-Z]+')

# FileMaker Pro database data
cnxn = pyodbc.connect('DSN=ParisAuctionNet;UID=emonson;PWD=test.log')
cursor = cnxn.cursor()

# Stopwords file
f = codecs.open(os.path.join(out_dir,'fr_painting_stopwords.txt'), 'r', 'utf-8')
# strip off newlines and blank lines
stopwords = [xx.rstrip('\n') for xx in f.readlines() if xx != '\n']
f.close()

cursor.execute('select Subject from Transactions')
rows = cursor.fetchall()
# looks like the original is in 'latin-1', but can use
# this to shift to utf-8 unicode strings
subjects = [x[0].decode('utf-8') for x in rows]

terms_list = []
terms_freq = []
doc_term_dicts = []

# Build up both 
# Treat all subjects as sentences. Not a problem for this simple regex tokenizer
for ii,sent in enumerate(subjects):
	tokens = nltk.wordpunct_tokenize(sent)
	
	# Filter out stopwords and puctuation
	good_tokens = [tt.lower() for tt in tokens if ((tt.lower() not in stopwords) and (not ncall_re.search(tt)))]
	# initialize entry to zero if not present
	counts_dict = defaultdict(int)
	
	for token in good_tokens:
		if token not in terms_list:
			terms_list.append(token)
			terms_freq.append(0)
		idx = terms_list.index(token)
		counts_dict[idx] += 1
		terms_freq[idx] += 1
	doc_term_dicts.append(counts_dict)
	
# Write out data files for Blei LDA

# Terms
out = codecs.open(os.path.join(out_dir,'fr_paintings_terms.txt'), 'w', 'utf-8')
for tt in terms_list:
	out.write(tt + u'\n')
out.close()

# Frequencies
out = open(os.path.join(out_dir,'fr_paintings_freq.dat'), 'w')
for ii,term_dict in enumerate(doc_term_dicts):

	out.write(str(len(term_dict)))
	
	for term,freq in term_dict.items():
		out.write(' ' + str(term) + ':' + str(freq))
	out.write('\n')

		