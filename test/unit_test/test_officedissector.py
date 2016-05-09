#!/usr/bin/env python

"""Unit tests for OfficeDissector"""

__author__ = 'Brandon Gordon'
__email__ = 'bgordon@grierforensics.com'

import sys
import os
import unittest
import StringIO

from lxml import etree

from officedissector.doc import Document
from officedissector.zip import ZipCRCError
from officedissector.part import Part


class PackageTest(unittest.TestCase):
    def setUp(self):
        sys.stdout = self.test_stdout = StringIO.StringIO()

    # DEV-01.1
    def testFileName(self):
        doc1 = Document('testdocs/test.docx')
        self.assertTrue(doc1.type == 'Word')
        doc2 = Document('testdocs/test.xlsx')
        self.assertTrue(doc2.type == 'Excel')
        doc3 = Document('testdocs/test.pptx')
        self.assertTrue(doc3.type == 'PowerPoint')

    # DEV-01.1
    def testFileMacro(self):
        doc1 = Document('testdocs/test.docx')
        self.assertFalse(doc1.is_macro_enabled)
        doc2 = Document('testdocs/test.docm')
        self.assertTrue(doc2.is_macro_enabled)

    # DEV-01.1
    def testFileTemplate(self):
        doc1 = Document('testdocs/test.docx')
        self.assertFalse(doc1.is_template)
        doc1 = Document('testdocs/test.dotx')
        self.assertTrue(doc1.is_template)

    def testIfFileExists(self):
        with self.assertRaises(IOError):
            Document('fakefile.docx')

    # DEV-01.2
    def testZipfileProperties(self):
        doc1 = Document('testdocs/test.docx')
        self.assertEquals(doc1.zip().namelist()[0], '[Content_Types].xml')
        self.assertEquals(doc1.zip().comment, '')
        self.assertEquals(doc1.zip().part_extract('[Content_Types].xml').read(10),
                          b'\x3C\x3F\x78\x6D\x6C\x20\x76\x65\x72\x73')
        self.assertEquals(len(doc1.zip().namelist()), 17)

        doc2 = Document('testdocs/testzipattrib.docx')
        self.assertEquals(doc2.zip().part_info(
            '[Content_Types].xml').file_size, 1818)
        self.assertEquals(doc2.zip().part_info(
            '[Content_Types].xml').compress_size, 406)
        self.assertEquals(doc2.zip().part_info('[Content_Types].xml').date_time,
                          (2013, 07, 03, 15, 22, 12))
        self.assertEquals(doc2.zip().part_info('[Content_Types].xml').comment,
                          '')

        with self.assertRaises(ZipCRCError):
            Document('testdocs/badcrc.docx').zip()

    # DEV-01.3
    def testPart(self):
        part1 = Part(Document('testdocs/test.docx'), '/[Content_Types].xml')
        self.assertEquals(part1.name, '/[Content_Types].xml')
        self.assertEquals(part1.stream().read(10),
                          b'\x3C\x3F\x78\x6D\x6C\x20\x76\x65\x72\x73')

    # DEV-01.4
    def testPartCollection(self):
        doc1 = Document('testdocs/test.docx')
        self.assertEquals(doc1.parts[0].name, '/[Content_Types].xml')
        self.assertEquals(doc1.parts[0].stream().read(10),
                          b'\x3C\x3F\x78\x6D\x6C\x20\x76\x65\x72\x73')
        self.assertEquals(doc1.parts[2].name, '/word/_rels/document.xml.rels')
        self.assertEquals(doc1.part_by_name['/[Content_Types].xml'].name,
                          '/[Content_Types].xml')
        self.assertEquals(
            doc1.part_by_name['/[Content_Types].xml'].stream().read(10),
            b'\x3C\x3F\x78\x6D\x6C\x20\x76\x65\x72\x73')

    # DEV-02
    def testPartXML(self):
        part1 = Document('testdocs/test.docx').part_by_name['/word/document.xml']
        self.assertEquals(
            list(part1.xml().getroot().iterchildren())[0].tag,
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}body')
        self.assertEquals(part1.xpath('//@w:val', part1.xml().getroot().nsmap)[2], 'Funotenzeichen')
        part2 = Document('testdocs/test.docx').part_by_name['/[Content_Types].xml']
        xmlns = {'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
        self.assertEqual(part2.xpath('/ct:Types/ct:Override/@PartName', xmlns)[0],
                         '/word/document.xml')

        part3 = Document('testdocs/testutf16.docx').part_by_name['/word/document.xml']
        self.assertEquals(
            part3.xml().xpath('//*[local-name() = "lang"]/@*[local-name() = "val"]')[0],
            'en-US')

        part4 = Document('testdocs/testascii.docx').part_by_name['/word/document.xml']
        self.assertEquals(
            part4.xml().xpath('//*[local-name() = "lang"]/@*[local-name() = "val"]')[0],
            'en-US')

        doc5 = Document('testdocs/macros-non-standard.xlsm')
        self.assertEqual(doc5.features.macros[0].name, '/xl/new_name.bin')

        part6 = Document('testdocs/non-standard-namespace.docx').part_by_name['/word/document.xml']
        self.assertEquals(part1.xpath('//@fake:val', part6.xml().getroot().nsmap)[2], 'Funotenzeichen')

    # DEV-03.1 and DEV-03.2
    def testContentTypes(self):
        doc1 = Document('testdocs/test.docx')
        part1 = doc1.part_by_name['/word/document.xml']
        self.assertEquals(
            part1.content_type(),
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml')
        part2 = Document('testdocs/test.docx').part_by_name['/customXml/item1.xml']
        self.assertEquals(part2.content_type(), 'application/xml')

        self.assertEqual(
            doc1.parts_by_content_type('application/vnd.ms-word.stylesWithEffects+xml')[0].name,
            '/word/stylesWithEffects.xml')
        self.assertEqual(doc1.parts_by_content_type_regex('footnotes')[0].name,
                         '/word/footnotes.xml')
        self.assertEqual(doc1.parts_by_content_type_regex('properties')[1].name,
                         '/docProps/app.xml')

    # DEV-03.3-7
    def testRelationships(self):
        doc1 = Document('testdocs/test.docx')
        self.assertEqual(len(doc1.relationships), 13)
        reltype = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
        self.assertEqual(doc1.relationships_dict[reltype][0].target,
                         'word/document.xml')
        self.assertEqual(
            doc1.find_relationships_by_type('metadata/core-properties')[0].source.name,
            'RootPart')
        self.assertEqual(doc1.find_relationships_by_type(
            'metadata/core-properties')[0].source.content_type(),
            '(virtual root part)')
        self.assertEqual(doc1.find_relationships_by_type('metadata/core-properties')[0].id,
                         'rId2')
        self.assertEqual(doc1.find_relationships_by_type('metadata/core-properties')[0].target,
                         'docProps/core.xml')
        self.assertEqual(doc1.find_relationships_by_type('metadata/core-properties')[0].target_part.name,
                         '/docProps/core.xml')
        self.assertEqual(doc1.find_relationships_by_type('metadata/core-properties')[0].type,
                         'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties')
        self.assertEqual(doc1.find_relationships_by_type('metadata/core-properties')[0].is_external,
                         False)

        doc2 = Document('testdocs/216688.docx')
        self.assertEqual(doc2.find_relationships_by_type('hyperlink')[0].is_external, True)

        self.assertEqual(doc1.part_by_name['/word/document.xml'].relationships_out()[2].id, 'rId7')
        self.assertEqual(doc1.part_by_name['/docProps/app.xml'].relationships_out(), [])
        self.assertEqual(doc1.part_by_name['/docProps/app.xml'].relationships_in()[0].id, 'rId3')

        self.assertEqual(doc1.root_part.relationships_out()[1].id, 'rId2')

        self.assertEqual(doc1.parts_by_relationship_type('extended-properties')[0].name,
                         '/docProps/app.xml')
        self.assertEqual(doc1.parts_by_relationship_type('ships/extended-properties')[0].name,
                         '/docProps/app.xml')

        self.assertEqual(doc1.main_part().name, '/word/document.xml')

    # DEV-04.1
    def testCoreProperties(self):
        doc1 = Document('testdocs/test.docx')
        self.assertEqual(doc1.core_properties.name, '/docProps/core.xml')
        self.assertEqual(doc1.core_properties.category, 'Auxiliary')
        self.assertEqual(doc1.core_properties.content_status, '')
        self.assertEqual(doc1.core_properties.created, '2010-10-21T08:54:00Z')
        self.assertEqual(doc1.core_properties.creator, 'Klaus-Peter Eckert')
        self.assertEqual(doc1.core_properties.description,
                         'Footnotes and endnotes in different sections')
        self.assertEqual(doc1.core_properties.identifier, '')
        self.assertEqual(doc1.core_properties.keywords,
                         'rainbow, color, colour, couleur')
        self.assertEqual(doc1.core_properties.language, '')
        self.assertEqual(doc1.core_properties.last_modified_by,
                         'Klaus-Peter Eckert')
        self.assertEqual(doc1.core_properties.last_printed, '')
        self.assertEqual(doc1.core_properties.modified, '2010-10-21T09:05:00Z')
        self.assertEqual(doc1.core_properties.revision, '4')
        self.assertEqual(doc1.core_properties.subject, '')
        self.assertEqual(doc1.core_properties.title, '')
        self.assertEqual(doc1.core_properties.version, '')
        doc2 = Document('testdocs/no_core_props.docx')
        self.assertEqual(doc2.core_properties.name, '')

    # DEV-04.2-5
    def testFeatures(self):
        doc1 = Document('testdocs/content.docx')
        self.assertEqual(doc1.features.custom_properties, [])
        self.assertEqual(len(doc1.features.images), 14)
        self.assertEqual(
            [i.content_type() for i in doc1.features.images if i.name == '/word/media/image1.png'],
            ['image/png'])
        self.assertEqual(
            sorted(doc1.features.images, key=lambda part: part.name)[0].content_type(),
            'image/png')
        self.assertEqual(len(doc1.features.videos), 0)
        self.assertEqual(len(doc1.features.fonts), 2)
        self.assertEqual(
            sorted(doc1.features.fonts, key=lambda part: part.name)[0].name,
            '/word/fonts/font1.odttf')

        doc2 = Document('testdocs/sounds.pptx')
        self.assertEqual(
            sorted(doc2.features.sounds, key=lambda part: part.name)[0].name,
            '/ppt/media/audio1.wav')
        self.assertEqual(
            sorted(doc2.features.sounds, key=lambda part: part.name)[0].content_type(),
            'audio/wav')

        doc3 = Document('testdocs/macros.xlsm')
        self.assertEqual(
            sorted(doc3.features.macros, key=lambda part: part.name)[0].name,
            '/xl/vbaProject.bin')
        self.assertEqual(
            sorted(doc3.features.embedded_controls, key=lambda part: part.name)[0].name,
            '/xl/activeX/activeX1.xml')

        doc4 = Document('testdocs/content2.docx')
        self.assertEqual(
            sorted(doc4.features.embedded_packages, key=lambda part: part.name)[2].name,
            '/word/embeddings/Microsoft_Excel-Arbeitsblatt3.xlsx')

        self.assertEqual(len(doc1.features.embedded_objects), 10)
        self.assertEqual(
            sorted(doc1.features.embedded_objects, key=lambda part: part.name)[2].name,
            '/word/embeddings/Microsoft_Office_PowerPoint_97-2003_Presentation7.ppt')

    # DEV-05
    def testExportJSON(self):
        doc1 = Document('testdocs/test.docx')
        self.assertEqual(doc1.part_by_name['/word/document.xml'].to_reference(),
                         'Part [/word/document.xml]')
        self.assertEqual(doc1.part_by_name['/word/document.xml'].relationships_out()[2].to_reference(),
                         'Relationship [rId7] (source Part [/word/document.xml])')
        self.assertEqual(doc1.part_by_name['/word/document.xml'].to_json()[0:30],
                         '{\n    "content-type": "applica')
        self.assertEqual(doc1.relationships[0].to_json()[0:32],
                         '{\n    "source": "Part [RootPart]')
        self.assertEqual(doc1.to_json()[0:20], '{\n    "document": [\n')
        self.assertEqual(doc1.to_json(include_stream=True)[285:325],
                         '     "stream_b64": "PD94bWwgdmVyc2lvbj0i')

    def testBugs(self):
        # Regression test for BUG OXPA-83
        # Make sure Target_Part='NULL', in this case a Relationship with
        # Type 'image', is handled properly
        doc1 = Document('testdocs/037027.pptx')
        for image in doc1.features.images:
            test = image.name

    def testAssertions(self):
        # TQA-02.4
        with self.assertRaises(KeyError):
            Document('testdocs/bad_extension.doc')
            self.assertEqual(self.test_stdout.getvalue(),
                             'File extension is not an OOXML file type')
        # Skip this test: The document doesn't follow the spec, but is still openable
        # with self.assertRaisesRegexp(AssertionError,
        #                             'content_type of Part is empty'):
        #    Document('testdocs/missing_content_type.docx')
        with self.assertRaises(KeyError):
            Document('testdocs/missing_part.docx')
            self.assertEqual(self.test_stdout.getvalue(),
                             'target_path is not a valid Part: /word/endnotes.xml')
        with self.assertRaises(KeyError):
            Document('testdocs/missing_rel_target.docx')
            self.assertEqual(self.test_stdout.getvalue(),
                             'target_path is not a valid Part: /')
        with self.assertRaises(etree.XMLSyntaxError):
            Document('testdocs/corrupt_xml.docx')
            self.assertEqual(self.test_stdout.getvalue(),
                             'part cannot be parsed successfully: Part [/[Content_Types].xml]')


def main():
    os.chdir(os.path.abspath(os.path.dirname(__file__)))
    suite = unittest.TestLoader().loadTestsFromTestCase(PackageTest)
    result = unittest.TextTestRunner(verbosity=2).run(suite)
    if not result.wasSuccessful():
        sys.exit(1)


if __name__ == '__main__':
    main()
