#!/usr/bin/env python

"""Basic functional tests to ensure MASTIFF example plugins run correctly."""

import os
import sys
import unittest
import subprocess
import hashlib
import glob


class PluginTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.mastiff_dir = sys.argv[1]
        global test_dir
        cls.test_dir = test_dir
        docs = ['testdocs/037027.pptx',
                'testdocs/content.docx',
                'testdocs/sounds.pptx',
                'testdocs/macros.xlsm',
                'testdocs/url.docx',
                'testdocs/test.docx',
                'testdocs/macros2.xlsm']
        print 'Running MASTIFF plugins on test documents...'
        for doc in docs:
            cls.run_mastiff(doc)

    # DEV-08.1
    def testMultimediaPlugin(self):
        types = ['*.jpeg', '*.png', '*.gif', '*.wmf', '*.wdp', '*.wav',
                 '*.odttf', '*.emf']
        doc1 = 'testdocs/037027.pptx'
        parts1 = sorted(self.getPartsFromDir(doc1, types))
        self.assertEqual(len(parts1), 62)
        part1_stream = open(parts1[0], 'rb').read(5)
        self.assertEqual(part1_stream, '\xFF\xD8\xFF\xE0\x00')

        doc2 = 'testdocs/content.docx'
        parts2 = sorted(self.getPartsFromDir(doc2, types))
        self.assertEqual(len(parts2), 16)
        self.assertEqual(hashlib.md5(open(parts2[15], 'rb').read()).hexdigest(),
                         '118c09f568beef74740e482a70ffa730')

        doc3 = 'testdocs/url.docx'
        parts3 = self.getPartsFromDir(doc3, types)
        self.assertEqual(len(parts3), 1)
        self.assertEqual(hashlib.md5(open(parts3[0], 'rb').read()).hexdigest(),
                         '77aa950e32e8f0b690d750a292d682e7')

        doc4 = 'testdocs/test.docx'
        self.assertEqual(len(self.getPartsFromDir(doc4, types)), 0)

    # DEV-08.2
    def testURLPlugin(self):
        doc1 = 'testdocs/037027.pptx'
        urls1 = self.getURLsFromFile(doc1)
        self.assertEqual(len(urls1), 22)
        self.assertEqual(urls1[0], 'http://www.crh.noaa.gov/glossary.php?word=RIDGE')
        self.assertEqual(urls1[21], 'http://www.srh.noaa.gov/lub/images/events/2007/20070120/balloon.jpg')

        doc2 = 'testdocs/url.docx'
        urls2 = self.getURLsFromFile(doc2)
        self.assertEqual(len(urls2), 1)
        self.assertEqual(urls2[0], 'http://www.ll.mit.edu')

        doc3 = 'testdocs/test.docx'
        urls3 = self.getURLsFromFile(doc3)
        self.assertEqual(len(urls3), 0)

    # DEV-08.3
    def testEmbeddedCodePlugin(self):
        types = ['*vba*.bin']
        doc1 = 'testdocs/macros.xlsm'
        parts1 = self.getPartsFromDir(doc1, types)
        self.assertEqual(len(parts1), 1)
        part1md5 = hashlib.md5(open(parts1[0], 'rb').read()).hexdigest()
        self.assertEqual(part1md5, '5218035425e1e9568cd49e782d0ce92e')

        doc2 = 'testdocs/037027.pptx'
        self.assertEqual(len(self.getPartsFromDir(doc2, types)), 0)

    @classmethod
    def run_mastiff(cls, doc):
        os.chdir(cls.mastiff_dir)
        try:
            p = subprocess.Popen(['mas.py',
                                  '-c', os.path.join(cls.mastiff_dir, 'mastiff.conf'),
                                  os.path.join(cls.test_dir, doc)],
                                 stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        except OSError as e:
            if e.errno == os.errno.ENOENT:
                print '\nmas.py cannot be found.'
                print 'Please ensure that MASTIFF is installed.'
                sys.exit(1)
            else:
                raise
        out, err = p.communicate()
        if 'Could not read any configuration files' in out:
            print '\n\n' + out
            print 'Most likely your argument does not point to'
            print 'a valid MASTIFF source directory.'
            sys.exit(1)
        os.chdir(cls.test_dir)

    def getPartsFromDir(self, doc, types):
        print 'mastiffparts %s' % self.mastiff_dir
        parts_path = os.path.join(self.mastiff_dir,
                                  'work/log',
                                  hashlib.md5(open(doc, 'rb').read()).hexdigest(),
                                  'parts')
        files_grabbed = []
        for files in types:
            files_grabbed.extend(glob.glob(os.path.join(parts_path, files)))
        return files_grabbed

    def getURLsFromFile(self, doc):
        url_filepath = os.path.join(self.mastiff_dir,
                                  'work/log',
                                  hashlib.md5(open(doc, 'rb').read()).hexdigest(),
                                  'urls.txt')
        url_file = open(url_filepath, 'r')
        url_file.readline()  # Skip first line header
        urls = []
        for u in url_file.readlines():
            u = u.strip()
            if u:
                urls.append(u)
        url_file.close()
        return urls

# For some odd reason, abspath() in unittest class method gives
# incorrect path, so use global variable instead.
test_dir = os.path.abspath(os.path.dirname(__file__))

def main():
    if len(sys.argv) != 2:
        print "\nUsage: " + sys.argv[0] + " MASTIFF_SOURCE_DIR\n\n"
        sys.exit(1)
    os.chdir(os.path.abspath(os.path.dirname(__file__)))
    suite = unittest.TestLoader().loadTestsFromTestCase(PluginTest)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    if not result.wasSuccessful():
        sys.exit(1)

if __name__ == '__main__':
    main()