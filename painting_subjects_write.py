#!/usr/bin/python
# -*- coding: utf-8 -*-

import numpy as N
import codecs
import os

out_dir = '/Users/emonson/Data/Victoria/HilaryPaintingSubjects'
os.chdir(out_dir)

tb = N.genfromtxt('/Users/emonson/Desktop/PaintSub2.tab',delimiter='\t',dtype=(int,float,N.object_))

subjects = [x for (a,b,x) in tb]
ids = [str(a).zfill(4) for (a,b,x) in tb]

for id,sub in zip(ids,subjects):
	out = codecs.open(id + '.txt', 'w', 'utf-8')
	out.write(sub.decode('latin-1'))
	out.close()
	