#!/usr/bin/env python

"""
Script to take html of Fraunhofer site as an argument
and download all OOXML files.
"""

import re
import sys
import urllib2


html = sys.argv[1]

host = 'http://www.document-interoperability.com'


pattern = re.compile(r'/c/document_library/get_file.*?\"')
pattern2 = re.compile(r'images/document_library/.*?</a>')

f = open(html, 'rb')
htmlcontent = f.read()
f.close()

filenames = []

# get names
for result in pattern2.findall(htmlcontent):
    filenames.append(result.split(' ', 1)[1].strip('</a>'))

filecount = 0
for result in pattern.findall(htmlcontent):
    path = result.strip('"')
    ext = re.search(r'name=.*?\..*', path).group()[-4:].strip('.')
    
    filename = filenames[filecount]
    
    file = urllib2.urlopen(host+path).read()
    
    f = open((filename+'.'+ext), 'wb')
    f.write(file)
    f.close()
    
    print filename + '.' + ext
    filecount += 1
print 'Total: %d' % filecount
