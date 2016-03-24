#!/usr/bin/env python

"""Basic smoke tests to ensure classes and methods run on entire corpus"""

__author__ = 'Brandon Gordon'
__email__ = 'bgordon@grierforensics.com'

import os
import sys
import codecs
import unittest

from officedissector.doc import Document
from officedissector.zip import ZipCRCError


class SmokeTest(unittest.TestCase):
    def testDocOpen(self):
        cur_dir = os.path.dirname(os.path.realpath(__file__))
        corpus_path = [os.path.join(cur_dir, 'govdocs'), os.path.join(cur_dir, 'fraunhoferlibrary')]
        files = []
        for dir_ in corpus_path:
            for f in os.listdir(dir_):
                file_ = os.path.join(dir_, f)
                if os.path.isfile(file_):
                    files.append(file_)

        log = codecs.open('smoke_tests.log', 'w', encoding='utf-8')
        errorlog = codecs.open('smoke_tests_error.log', 'w', encoding='utf-8')

        # Write error log header
        errorlog.write('Three files from govdocs (govdocs/641559.docx, govdocs/500968.xlsx, and\n'
                       'govdocs/974690.xlsx) are apparently corrupt and do not open in Microsoft\n'
                       'Office 2010; hence, error messages should appear for them. (Note that govdocs\n'
                       'is a random sample and explicitly includes bad or corrupt files.)\n'
                       'No other error messages should appear here.\n\n\n')

        for docfile in files:
            print('\nTesting %s...' % docfile)
            log.write('\nTesting %s...\n' % docfile)
            try:
                doc1 = Document(docfile)
            except ZipCRCError:
                msg = 'Error: Bad CRC for file: %s\n' % docfile
                print(msg)
                log.write(msg)
                errorlog.write(msg)
                continue
            except Exception as e:
                msg = 'Error: File: %s: %s - %s\n' % (docfile, sys.exc_info()[0].__name__, e)
                print(msg)
                log.write(msg)
                errorlog.write(msg)
                continue

            log.write('  Document type is: %s\n' % doc1.type)
            log.write('  Document is macro enabled: %s\n' % doc1.is_macro_enabled)
            log.write('  Document is a template: %s\n' % doc1.is_template)

            print('  Testing zip.part_info method...')
            log.write('  Testing zip.part_info method...\n')
            log.write('    zip.part_info([Content_Types].xml).file_size: %s\n' %
                      doc1.zip().part_info('[Content_Types].xml').file_size)
            log.write('    zip.part_info([Content_Types].xml).compress_size: %s\n' %
                      doc1.zip().part_info('[Content_Types].xml').compress_size)
            print('  Done.')
            log.write('  Done.\n')

            second_part = doc1.parts[1]
            print('  Testing Part: %s' % second_part.name)
            log.write('  Testing Part: %s\n' % second_part.name)

            doc_stream = doc1.part_by_name[second_part.name].stream().read(10)
            print('  Part stream successfully captured.')
            log.write('  Part stream successfully captured.\n')
            partxml = doc1.part_by_name[second_part.name].xml()
            print('  Part XML successfully parsed.')
            log.write('  Part XML successfully parsed.\n')

            print('  Checking doc.xpath method...')
            log.write('  Checking doc.xpath method...\n')
            log.write('    XPath Result: %s\n' %
                      doc1.part_by_name['/[Content_Types].xml'].xpath('*/@ContentType')[0])
            print('  Done.')
            log.write('  Done.\n')

            print('  Checking that all Parts can get Content_Type...')
            log.write('  Checking that all Parts can get Content_Type...\n')
            for part in doc1.parts:
                ct = part.content_type()
                log.write('    Part %s is Content_Type: %s\n' % (part.name, ct))
            print('  Done.')
            log.write('  Done.\n')

            print('  Checking that Document has main_part...')
            log.write('  Checking that Document has main_part...\n')
            doc_main = doc1.main_part()
            log.write('    Main Part: %s\n' % doc_main.name)
            print('  Done.')
            log.write('  Done.\n')

            print('  Checking all source and target Relationships for each part...')
            log.write('  Checking all source and target Relationships for each part...\n')
            for part in doc1.parts:
                rel_in = part.relationships_in()
                rel_out = part.relationships_out()
                log.write('    Part %s: Relationships in source name: %s\n' %
                          (part.name, [r.source.name for r in rel_in]))
                log.write('    Part %s: Relationships out: %s\n' %
                          (part.name, [r.target for r in rel_out]))
            print('  Done.')
            log.write('  Done.\n')

            print('  Testing Document methods to find by Part or Relationship...')
            log.write('  Testing Document methods to find by Part or Relationship...\n')
            log.write('    doc.parts_by_content_type(application/xml): %s\n' %
                      doc1.parts_by_content_type('application/xml')[0])
            log.write('    doc.parts_by_content_type_regex(ation/xm): %s\n' %
                      doc1.parts_by_content_type_regex('ation/xm')[0])
            log.write('    doc.parts_by_relationship_type(/relationships/officeDocument: %s\n' %
                      doc1.parts_by_relationship_type('/relationships/officeDocument')[0].name)
            log.write('    doc.find_relationship_by_type(/relationships/officeDocument).source: %s\n' %
                      doc1.find_relationships_by_type('/relationships/officeDocument')[0].source)
            print('  Done.')
            log.write('  Done.\n')

            print('  Checking for all Features...')
            log.write('  Checking for Features...\n')
            for image in doc1.features.images:
                log.write('    Image: %s\n' % image.name)
            for video in doc1.features.videos:
                log.write('    Video: %s\n' % video.name)
            for sound in doc1.features.sounds:
                log.write('    Sound: %s\n' % sound.name)
            for font in doc1.features.fonts:
                log.write('    Font: %s\n' % font.name)
            for macro in doc1.features.macros:
                log.write('    Macro content: %s\n' % macro.name)
            for comment in doc1.features.comments:
                log.write('    Comment content: %s\n' % comment.name)
            for customX in doc1.features.custom_xml:
                log.write('    Custom XML content: %s\n' % customX.name)
            for embedded_control in doc1.features.embedded_controls:
                log.write('    Embedded Control content: %s\n' % embedded_control.name)
            for embedded_object in doc1.features.embedded_objects:
                log.write('    Embedded Object content: %s\n' % embedded_object.name)
            for embedded_package in doc1.features.embedded_packages:
                log.write('    Embedded Package content: %s\n' % embedded_package.name)
            for digital_signature in doc1.features.digital_signatures:
                log.write('    Digital Signature content: %s\n' % digital_signature.name)
            print('  Done.')
            log.write('  Done.\n')

            print('  Checking Core Properties...')
            log.write('  Checking Core Properties...\n')
            log.write('    Category: %s\n' % doc1.core_properties.category)
            log.write('    Content status: %s\n' % doc1.core_properties.content_status)
            log.write('    Created: %s\n' % doc1.core_properties.created)
            log.write('    Creator: %s\n' % doc1.core_properties.creator)
            log.write('    Description: %s\n' % doc1.core_properties.description)
            log.write('    Identifier: %s\n' % doc1.core_properties.identifier)
            log.write('    Keywords: %s\n' % doc1.core_properties.keywords)
            log.write('    Language: %s\n' % doc1.core_properties.language)
            log.write('    Last modified by: %s\n' % doc1.core_properties.last_modified_by)
            log.write('    Last printed: %s\n' % doc1.core_properties.last_printed)
            log.write('    Modified: %s\n' % doc1.core_properties.modified)
            log.write('    Revision: %s\n' % doc1.core_properties.revision)
            log.write('    Subject: %s\n' % doc1.core_properties.subject)
            log.write('    Title: %s\n' % doc1.core_properties.title)
            log.write('    Version: %s\n' % doc1.core_properties.version)
            print('  Done.')
            log.write('  Done.\n')

            print('  Checking export to JSON...')
            log.write('  Checking export to JSON...\n')
            doc_json = doc1.to_json()
            log.write('    Beginning of JSON: %s\n' % doc_json[0:50])
            print('  Done.')
            log.write('  Done.\n')

            print('Done.')
            log.write('Done.\n')

        log.close()
        errorlog.close()


def main():
    os.chdir(os.path.abspath(os.path.dirname(__file__)))
    suite = unittest.TestLoader().loadTestsFromTestCase(SmokeTest)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    if not result.wasSuccessful():
            sys.exit(1)


if __name__ == '__main__':
    main()
