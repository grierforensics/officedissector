#!/usr/bin/env python

"""A Part of a Document."""

__author__ = 'Brandon Gordon'
__email__ = 'bgordon@grierforensics.com'

import json
import base64
import StringIO

from lxml import etree
from types import *


class Part(object):
    """
    A :class:`Part` of a Document.

    :ivar doc: the Document object associated with this Part.
    :ivar name: name of :class:`Part`.
        Always has a preceding '/' eg. '/word/document.xml'
    """

    def __init__(self, doc, name):
        """
        Initialize the :class:`Part`.

        :param doc: the :class:`~officedissector.doc.Document` object associated with this :class:`Part`
        :type doc: :class:`~officedissector.doc.Document`
        :param name: the name of the :class:`Part`
        :type name: `string`
        """
        self.name = name
        self.doc = doc
        self.__content_type = None

    def stream(self):
        """
        Return a file-like object of this :class:`Part`.

        :return: a file-like object"""
        stream_ = self.doc.zip().part_extract(self.name)
        assert stream_ is not None, 'stream is empty: %r' % stream_
        return stream_

    def xml(self):
        """
        Parse the XML of :class:`Part`.

        :return: an `ElementTree` of the parsed XML
        """
        try:
            xml_etree = etree.parse(self.stream())
        except etree.XMLSyntaxError:
            print 'part cannot be parsed successfully: %r' % self
            raise
        return xml_etree

    def xpath(self, exp, xmlns=None):
        """
        Evaluate the XPath expression 'exp' on the XML of the :class:`Part`.

        To properly understand how to use xpath, see the documentation
        of XML :doc:`Namespaces`.

        The suggested method to use xpath is to provide a user-defined
        namespace mapping. For example:

        >>> CP_NAMESPACE = {'cp': "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        >>>                 'dc': "http://purl.org/dc/elements/1.1/",
        >>>                 'dcterms': "http://purl.org/dc/terms/",
        >>>                 'dcmitype': "http://purl.org/dc/dcmitype/",
        >>>                 'xsi': "http://www.w3.org/2001/XMLSchema-instance"}
        >>> result = part.xpath('/cp:coreProperties/dcterms:created', CP_NAMESPACE)
        >>> result[0].text
        '2008-11-03T19:31:00Z'


        :param exp: the XPATH expression
        :type exp: `string`
        :param xmlns: Optional - the namespace mapping (default: `None`)
        :type xmlns: `dict`
        :return: the result of the XPATH query.
        """
        xmletree = self.xml()

        return xmletree.xpath(exp, namespaces=xmlns)

    def content_type(self):
        """
        Determine Content Type of this :class:`Part`
        by parsing [Content_Types].xml

        For information about how Content Types are stored, see:
        http://office.microsoft.com/en-us/office-open-xml-i-exploring-the-office-open-xml-formats-RZ010243529.aspx?section=16

        :return: the Content Type of this :class:`Part`
        """
        if self.__content_type is not None:
            return self.__content_type

        content_types = self.doc.part_by_name['/[Content_Types].xml']

        # Since /[Content_Types].xml uses the default namespace (no prefix),
        # we define our own prefix: 'ct' to use in the XPath expression.
        ct_nsmap = {'ct': 'http://schemas.openxmlformats.org/package/2006/content-types'}
        # XPath query if entire Part is referenced in [Content_Type].xml
        xpath_qry1 = '/ct:Types/ct:Override[@PartName="' + self.name + '"]/@ContentType'
        result1 = content_types.xpath(xpath_qry1, ct_nsmap)
        if len(result1):
            self.__content_type = result1[0]
        else:
            # If the Part name is not in Override, get
            # ContentType based on extension of the Part name
            xpath_qry2 = ('/ct:Types/ct:Default[@Extension="' +
                          self.name.rsplit('.', 1)[1].strip('.') +
                          '"]/@ContentType')
            result2 = content_types.xpath(xpath_qry2, ct_nsmap)
            # When this second XPath result is also empty, this Part has no content_type
            assert len(result2), 'content_type of Part is empty: %r' % self
            self.__content_type = result2[0]

        return self.__content_type

    def relationships_out(self):
        """
        Determine all :class:`Relationship` objects
        for which this part is the source.

        >>> for rel in doc.part_by_name['/word/document.xml'].relationships_out():
        >>>     print rel.source
        Part [/word/document.xml]
        Part [/word/document.xml]
        Part [/word/document.xml]

        :return: list of all Relationships out
        """
        rels_out = []
        for rel in self.doc.relationships:
            if rel.source == self:
                rels_out.append(rel)
        return rels_out

    def relationships_in(self):
        """
        Determine all :class:`Relationship` objects
        for which this part is the target.

        >>> doc.part_by_name['/word/document.xml'].relationships_in()[0].target
        'word/document.xml'

        :return: list of all Relationships in
        """
        rels_in = []
        for rel in self.doc.relationships:
            if rel.target_part == self:
                rels_in.append(rel)
        return rels_in

    def to_reference(self):
        """
        A which uniquely identifies this object.

        :return: string which uniquely identifies this object"""
        return "Part [%s]" % self.name

    def to_json(self, include_stream=False):
        """
        Export this object to JSON

        :param include_stream: Optional - Include base64 encoded stream of
            this :class:`Part` (Default false).
        :type include_stream: `bool`
        :return: a JSON encoded string"""
        rels_in = []
        for rel_in in self.relationships_in():
            rels_in.append(rel_in.to_reference())
        rels_out = []
        for rel_out in self.relationships_out():
            rels_out.append(rel_out.to_reference())

        json_dump = {'uri': self.name, 'content-type': self.content_type(),
                     'relationships_in': rels_in, 'relationships_out': rels_out}

        if include_stream:
            stream_encoded = base64.b64encode(self.stream().read())
            json_dump['stream_b64'] = stream_encoded

        json_str = json.dumps(json_dump, indent=4)
        try:
            json.loads(json_str)
        except ValueError:
            print 'json_str is not valid JSON'
            raise

        return json_str

    def __repr__(self):
        return self.to_reference()


class RootPart(Part):
    """
    A special part object, used when the source of a
    :class:`Relationship` object is the virtual root Part.
    """

    def __init__(self, doc):
        self.name = 'RootPart'
        self.doc = doc

    def stream(self):
        """
        In the RootPart, for the stream method, return an empty stream.

        :return: empty stream
        """
        return StringIO.StringIO()

    def xml(self):
        """
        In the RootPart, for the xml method, return `None`.

        return: `None`
        """
        return None

    def xpath(self, exp, xmlns=None):
        """
        In the RootPart, for the xpath method, return `None`.

        return: `None`
        """

    def content_type(self):
        """
        For the RootPart, return (virtual root part) as content_type.

        return: '(virtual root part)'
        """
        return '(virtual root part)'