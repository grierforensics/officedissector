#!/usr/bin/env python

"""Parses and exposes the Core Properties of the Document."""

__author__ = 'Brandon Gordon'
__email__ = 'bgordon@grierforensics.com'


class CoreProperties(object):
    """
    Parse and expose the Core Properties of the :class:`~officedissector.doc.Document`.

    Parse the Core Properties with element names and namespace prefixes
    according to the OOXML specifications in `ISO/IEC:29500-2.11 <http://standards.iso.org/ittf/PubliclyAvailableStandards/c061796_ISO_IEC_29500-2_2012_Electronic_inserts.zip>`_.

    :ivar category:
    :ivar content_status:
    :ivar created:
    :ivar creator:
    :ivar description:
    :ivar identifier:
    :ivar language:
    :ivar last_modified_by:
    :ivar last_printed:
    :ivar modified:
    :ivar revision:
    :ivar subject:
    :ivar title:
    :ivar version:

    """
    def __init__(self, core_prop_part):
        """
        Initialize the Core Properties object.

        :param core_prop_part: the :class:`~officedissector.part.Part` which contains the
            information about Core Properties.
        :type core_prop_part: :class:`~officedissector.part.Part`
        """
        if core_prop_part is None:
            self.name = ''
        else:
            self.core_prop_part = core_prop_part
            self.name = core_prop_part.name

        self.category = ''
        self.content_status = ''
        self.created = ''
        self.creator = ''
        self.description = ''
        self.identifier = ''
        self.language = ''
        self.last_modified_by = ''
        self.last_printed = ''
        self.modified = ''
        self.revision = ''
        self.subject = ''
        self.title = ''
        self.version = ''
        
    def parse_all(self):
        """Parse all Core Properties."""
        self.category = self._parse_prop('/cp:coreProperties/cp:category')
        self.content_status = self._parse_prop('/cp:coreProperties/cp:contentStatus')
        self.created = self._parse_prop('/cp:coreProperties/dcterms:created')
        self.creator = self._parse_prop('/cp:coreProperties/dc:creator')
        self.description = self._parse_prop('/cp:coreProperties/dc:description')
        self.identifier = self._parse_prop('/cp:coreProperties/dc:identifier')
        self.language = self._parse_prop('/cp:coreProperties/dc:language')
        self.last_modified_by = self._parse_prop('/cp:coreProperties/cp:lastModifiedBy')
        self.last_printed = self._parse_prop('/cp:coreProperties/cp:lastPrinted')
        self.modified = self._parse_prop('/cp:coreProperties/dcterms:modified')
        self.revision = self._parse_prop('/cp:coreProperties/cp:revision')
        self.subject = self._parse_prop('/cp:coreProperties/dc:subject')
        self.title = self._parse_prop('/cp:coreProperties/dc:title')
        self.version = self._parse_prop('/cp:coreProperties/cp:version')

        # Special parsing for Keywords, which can have subelements
        keywords_prop = self.core_prop_part.xpath('/cp:coreProperties/cp:keywords', CP_NAMESPACE)
        self.keywords = [k.text for k in keywords_prop if k.text is not None]
        if len(keywords_prop) > 0:
            added_keywords = [', '.join([k.text for k in keywords_prop[0]])]
            self.keywords = ', '.join(self.keywords + added_keywords)

    def _parse_prop(self, exp):
        """Return property if exists, if not, returns empty string"""
        result = self.core_prop_part.xpath(exp, CP_NAMESPACE)
        if result:
            return result[0].text
        else:
            return ''

    def __repr__(self):
        return "Core Properties of: %s" % self.core_prop_part.doc

# Namespace prefixes and URIs
# Source: ISO/IEC:29500-2.11
CP_NAMESPACE = {'cp': "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                'dc': "http://purl.org/dc/elements/1.1/",
                'dcterms': "http://purl.org/dc/terms/",
                'dcmitype': "http://purl.org/dc/dcmitype/",
                'xsi': "http://www.w3.org/2001/XMLSchema-instance"}