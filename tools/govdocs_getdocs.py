#!/usr/bin/env python

"""Script to crawl Govdocs index pages and get all OOXML docs."""

import urllib2
import re

url = "http://digitalcorpora.org/"
path = "corp/nps/files/govdocs1/"
ext = ['docx', 'docm', 'xlsx', 'xlsm', 'pptx', 'pptm']
pattern = re.compile('.......%s</A>' % ext, re.IGNORECASE)

for i in range(2):
    dir = '%03d' % i
    response = urllib2.urlopen(url+path+dir).read()
    
    for e in ext:
        pattern = re.compile('.......%s</A>' % e, re.IGNORECASE)
        for match in pattern.findall(response):
            filename = match.replace('</a>', '')
            print filename
            f = open(filename, 'wb')
            file = urllib2.urlopen(url+path+dir+'/'+filename).read()
            f.write(file)
            


