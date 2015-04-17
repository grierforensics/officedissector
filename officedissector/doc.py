#!/usr/bin/env python

"""An OOXML Document."""

__author__ = 'Brandon Gordon'
__email__ = 'bgordon@grierforensics.com'

import os
import posixpath
import re
import json

from collections import defaultdict

from officedissector.zip import Zip
from officedissector.part import Part
from officedissector.part import RootPart
from officedissector.rel import Relationship
from officedissector.core_properties import CoreProperties
from officedissector.features import Features


class Document(object):
    """
    A OOXML document.

    :ivar filepath: The path to the OOXML document.

    :ivar type: The OOXML document type, eg. Word.

    :ivar is_macro_enabled: True if document is macro enabled.

    :ivar is_template: True if document is a template.

    :ivar parts: List of all Parts in Document

    :ivar parts_by_name: Dictionary of all Parts with name as the key.

    :ivar root_part: Singleton of the RootPart class.
        Used to represent the virtual root part as the source of a Relationship.

    :ivar relationships: List of Relationships in the Document.

    :ivar relationships_dict: Dictionary of all Relationships
        with the full Relationship Type as the key.

    :ivar features: Object which contains all Features of the Document

    :ivar core_properties: Object which contains all Core Properties of the Document.

    """

    def __init__(self, filepath):
        """
        Initialize attributes. Build collections of Parts
        and Relationships.
        Parse and provide access to Features and CoreProperties.

        :param filepath: the path to the document file
        :type filepath: string
        """
        self.filepath = filepath

        # Does file exists?
        with open(filepath):
            pass

        # Is file's zip CRC is correct?
        self.zip().testzip()

        filename, ext = os.path.splitext(filepath)
        try:
            self.type, self.is_macro_enabled, self.is_template = FILE_EXTS[ext]
        except KeyError:
            print 'File extension is not an OOXML file type: %s' % ext
            raise

        # Provide list and dictionary of all Parts in :class:`Document`
        self.parts = []
        self.part_by_name = {}
        for name in self.zip().namelist():
            if name.endswith('/'):  # Skip directories of zip file
                continue

            name = '/' + name
            newpart = Part(self, name)
            self.parts.append(newpart)
            self.part_by_name[name] = newpart

        # Instantiate Singleton RootPart class
        self.root_part = RootPart(self)

        self.relationships, self.relationships_dict = self._parse_relationships()

        self.features = Features(self)

        self.core_properties = self._parse_core_properties()

    def zip(self):
        """
        Return a Zip object of OOXML located at self.filepath.

        :return: Zip object
        """
        return Zip(self.filepath)

    def parts_by_content_type(self, contype):
        """
        Determines list of all Parts with a given content-type.

        :param contype: content-type
        :type contype: string
        :return: list of all Parts with content-type
        """
        partslist = []
        for part in self.parts:
            if part.content_type() == contype:
                partslist.append(part)
        return partslist

    def parts_by_content_type_regex(self, exp):
        """
        Determine list of all Parts with content-type
        matching regex.

        :param exp: regular expression to match content-type of parts
        :type exp: string
        :return: list of all Parts with content-type matching regex
        """
        partslist = []
        for part in self.parts:
            if re.search(exp, part.content_type()):
                partslist.append(part)
        return partslist

    def parts_by_relationship_type(self, reltype):
        """
        Determine list of all Parts with incoming Relationships where
        reltype matches the end of the Relationship Type.

        For example:

        >>> part = doc.parts_by_relationship_type('ships/officeDocument')
        >>> print part[0]
        [Part [/word/document.xml]]
        >>> part[0].relationships_in()[0].type
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'.

        :param reltype: :class:`~officedissector.rel.Relationship` type. Match the end of reltype.
        :type reltype: string
        :return: list of all Parts with reltype matching the
            end of the Relationship type
        """
        part_list = []
        # Match the end of the Relationship Type.
        regex_type = reltype + '$'
        for rel in self.relationships:
            if (re.search(regex_type, rel.type) and (not rel.is_external)
                    and (rel.target_part is not None)):
                part_list.append(rel.target_part)
        return part_list

    def main_part(self):
        """
        Get the main document :class:`~officedissector.part.Part`.

        :return: the main document :class:`~officedissector.part.Part`
        """
        parts = self.parts_by_relationship_type('officeDocument')
        if len(parts) == 1:
            return parts[0]
        else:
            raise Exception("Document %s has %s main parts defined; it should have exactly 1" %
                            (self.filepath, len(parts)))

    def find_relationships_by_type(self, reltype):
        """
        Determine list of all Relationships where reltype
        matches the end of the Relationship Type.

        For example:

        >>> rel = doc.find_relationship_by_type('ships/officeDocument')
        >>> rel[0].type
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

        :param reltype: :class:`~officedissector.rel.Relationship` type. Match the end of reltype.
        :type reltype: string
        :return: list of all Relationships with reltype matching the
            end of reltype
        """
        regex_type = reltype + '$'
        rel_list = []
        for rel in self.relationships:
            if re.search(regex_type, rel.type):
                rel_list.append(rel)
        return rel_list

    def to_json(self, include_stream=False):
        """
        Export this object to JSON

        :param include_stream: Optional - Include base64 encoded stream of
            all Parts (Default false).
        :type include_stream: bool
        :return: a JSON encoded string
        """
        parts_json = []
        for part in self.parts:
            parts_json.append(json.loads(part.to_json(include_stream)))
        rels_json = []
        for rel in self.relationships:
            rels_json.append(json.loads(rel.to_json()))

        json_str = json.dumps({'document': [{'parts': parts_json}, {'relationships': rels_json}]}, indent=4)
        try:
            json.loads(json_str)
        except ValueError:
            print 'json_str is not valid JSON'
            raise
        return json_str

    def __repr__(self):
        return "Document: %s" % self.filepath

    def _parse_relationships(self):
        """
        Parse all .rels parts and create a Relationship object for each relationship.

        :return: list and dictionary of Relationships.
        """
        relationships = []

        # A defaultdict(list) is needed so multiple relationships
        # with the same type, and therefore the same key in the dict,
        # can be appended as a list of Relationships to a single dict entry.
        relationships_dict = defaultdict(list)

        # Since .rels parts use the default namespace,
        # define our own prefix 'rel' for the default namespace
        # to use in the XPath expression.
        relnsmap = {'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'}

        for relpart in self.parts_by_content_type('application/vnd.openxmlformats-package.relationships+xml'):

            for rel in relpart.xpath('/rel:Relationships/rel:Relationship',
                                     relnsmap):
                # Determine source by ignoring the '.rels' extension
                sourcename = relpart.name.rsplit('.', 1)[0]
                # Build the source path by removing the '_rels' directory from
                # the path.
                sourcepath = sourcename.rsplit(
                    '/', 2)[0] + '/' + sourcename.rsplit('/', 2)[2]

                if sourcepath == '/':
                    source = self.root_part
                else:
                    try:
                        source = self.part_by_name[sourcepath]
                    except KeyError:
                        print 'sourcepath is not a valid Part: %s' % sourcepath
                        raise

                reltype = rel.attrib['Type']
                relid = rel.attrib['Id']
                target = rel.attrib['Target']

                if 'TargetMode' in rel.attrib:
                    is_external = True if rel.attrib['TargetMode'] == 'External' else False
                else:
                    is_external = False

                if is_external:
                    target_part = None
                # If Target='NULL', looks like dangling relationship, so we
                # don't have a target part.
                # From the smoke tests, we found the following file
                # has these 'NULL' Targets: /govdocs/037027.pptx
                # See http://answers.microsoft.com/en-us/office/forum/office_2007-word/the-image-part-with-relationship-id-rid308-was-not/6bf26696-c7e8-49e3-8808-34b4fe40b1ec
                elif target == 'NULL':
                    target_part = None
                else:
                    # Build complete target_part path: begin with source name,
                    # go up 2 dirs, add target name.
                    # If Target has relative path: '../customXml/item1.xml';
                    # posixpath.normpath normalizes a Unix style path.
                    target_path = posixpath.normpath(sourcename.rsplit('/', 2)[0] + '/' + target)
                    try:
                        target_part = self.part_by_name[target_path]
                    except KeyError:
                        print 'target_path is not a valid Part: %s' % target_path
                        raise

                newrelobj = Relationship(source, reltype, relid, target, target_part, is_external)
                relationships.append(newrelobj)
                relationships_dict[reltype].append(newrelobj)
        return relationships, relationships_dict

    def _parse_core_properties(self):
        """
        Identify, extract and parse the Core Properties of the Document.

        :return: Core Properties object.
        """
        core_props1 = self.parts_by_content_type('application/vnd.openxmlformats-package.core-properties+xml')
        core_props2 = self.parts_by_relationship_type('metadata/core-properties')
        # We retrieve Core Properties based on both Content Type
        # and Relationship Type. Usually, each will return the
        # same Part. set() removes duplicate Parts.
        core_props = set(core_props1 + core_props2)
        if len(core_props) > 0:
            assert len(core_props) == 1, 'more than one core_properties Part: %s' % \
                                         [part.name for part in core_props]
            core_properties = CoreProperties(core_props.pop())
            core_properties.parse_all()
        else:
            core_properties = CoreProperties(None)
        return core_properties

# OOXML Attributes by File Extension
# Schema: {extension: (type, macro_enabled, is_template)}
# Source: http://office.microsoft.com/en-us/powerpoint-help/introduction-to-new-file-name-extensions-HA010006935.aspx?CTT=1
FILE_EXTS = {
    '.docx': ('Word', False, False),
    '.docm': ('Word', True, False),
    '.dotx': ('Word', False, True),
    '.dotm': ('Word', True, True),
    '.xlsx': ('Excel', False, False),
    '.xlsm': ('Excel', True, False),
    '.xltx': ('Excel', False, True),
    '.xltm': ('Excel', True, True),
    '.xlsb': ('Excel binary', False, False),
    '.xlam': ('Excel add-in', True, False),
    '.pptx': ('PowerPoint', False, False),
    '.pptm': ('PowerPoint', True, False),
    '.potx': ('PowerPoint', False, True),
    '.potm': ('PowerPoint', True, True),
    '.ppam': ('PowerPoint add-in', True, False),
    '.ppsx': ('PowerPoint show', False, False),
    '.ppsm': ('PowerPoint show', True, False),
    '.sldx': ('PowerPoint slide', False, False),
    '.sldm': ('PowerPoint slide', True, False),
    '.thmx': ('Office theme', False, False)
    }


