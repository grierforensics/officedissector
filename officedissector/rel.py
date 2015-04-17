#!/usr/bin/env python

"""A Relationship in an OOXML Document."""

__author__ = 'Brandon Gordon'
__email__ = 'bgordon@grierforensics.com'

import json


class Relationship(object):
    """
    A :class:`Relationship` in an OOXML Document.

    :ivar source: The source of a Relationship.
        For example, a Relationship in the
        '/word/_rels/document.xml.rels' Part has the
        source: '/word/document.xml'. For Relationships in the
        '/_rels/.rels' Part, the source is the virtual
        root part, or RootPart.

    :ivar type: The Type of a Relationship. For example, the
        Relationship with the Target '/word/document.xml'
        is Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

    :ivar id: The unique ID for this Relationship object. Note that the ID is
        unique only in that particular .rels file. The ID may be reused in
        another .rels file.

    :ivar target: The URI of the relationship's target (e.g. footnotes.xml)

    :ivar target_part: A reference to the target Part object.

    :ivar is_external: True if the Target of the Relationship refers to
        an external resource (eg. a hyperlink).
    """

    def __init__(self, source, type, id, target, target_part, is_external):
        self.source = source
        self.type = type
        self.id = id
        self.target = target
        self.target_part = target_part
        self.is_external = is_external

    def to_reference(self):
        """
        Return a string which uniquely identifies this object.

        :return: a string which uniquely identifies this object
        """
        return "Relationship [%s] (source %s)" % \
               (self.id, self.source.to_reference())

    def to_json(self):
        """
        Export this object to JSON.

        :return: a JSON encoded string
        """
        json_str = json.dumps(
            {'source': self.source.to_reference(), 'target': self.target,
             'type': self.type, 'id': self.id,
             'is_external': self.is_external}, indent=4)
        try:
            json.loads(json_str)
        except ValueError:
            print 'json_str is not valid JSON'
            raise
        return json_str

    def __repr__(self):
            return self.to_reference()