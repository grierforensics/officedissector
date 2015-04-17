#!/usr/bin/env python

"""Provides access to all Features of a Document."""

__author__ = 'Brandon Gordon'
__email__ = 'bgordon@grierforensics.com'


class Features(object):
    """
    Provide access to all :class:`Features` of a :class:`~officedissector.doc.Document`.

    Features in an OOXML document are identified in two ways:

    1. By their Content-Type
    2. By their inbound Relationship types

    For completeness, we use either means of identification,
    allowing either one to identify a feature.

    For example:

    >>> part = doc.part_by_name['/word/media/image1.jpeg']
    >>> part.content_type()
    'image/jpeg'
    >>> part.relationships_in()[0].type
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

    Content-Types and Relationships are mainly referenced from the OOXML
    specifications at `ISO/IEC:29500-1`_ 15.2 and `ISO/IEC:29500-2`_ 13.2.

    .. _`ISO/IEC:29500-2`:
        http://standards.iso.org/ittf/PubliclyAvailableStandards/c061796_ISO_IEC_29500-2_2012_Electronic_inserts.zip
    .. _`ISO/IEC:29500-1`:
        http://standards.iso.org/ittf/PubliclyAvailableStandards/c061750_ISO_IEC_29500-1_2012.zip

    :ivar custom_properties:
    :ivar images:
    :ivar videos:
    :ivar sounds:
    :ivar fonts:
    :ivar macros:
    :ivar comments:
    :ivar custom_xml:
    :ivar embedded_controls:
    :ivar embedded_objects:
    :ivar embedded_packages:
    :ivar digital_signatures:

    """

    def __init__(self, doc):
        """
        Initialize the Features object.

        :param doc: the :class:`~officedissector.doc.Document` associated with this object
        :type doc: :class:`~officedissector.doc.Document`
        """
        self.doc = doc

        # Schema: self._get_parts([content_type1,
        #                          content_type2...],
        #                         [relationships1,
        #                          relationship2...])
        self.custom_properties = self._get_parts(
            ['application/vnd.openxmlformats-officedocument.custom-properties+xml'],
            ['custom-properties'])

        self.images = self._get_parts(['image/'],
                                      ['relationships/image'])

        self.videos = self._get_parts(['video/'],
                                      ['relationships/video'])

        self.sounds = self._get_parts(['audio/'],
                                      ['relationships/audio'])

        self.fonts = self._get_parts(
            ['application/x-font',
             'application/vnd.openxmlformats-officedocument.obfuscatedFont'],
            ['relationships/font'])

        self.macros = self._get_parts(['application/vnd.ms-office.vbaProject',
                                       'application/vnd.ms-excel.intlmacrosheet+xml'],
                                      ['relationships/xlIntlMacrosheet',
                                       'relationships/xlIntlMacrosheet',
                                       'relationships/vbaProject'])

        self.comments = self._get_parts(
            ['application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml',
             'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
             'application/vnd.openxmlformats-officedocument.presentationml.comments+xml'],
            ['relationships/comments'])

        # Note that Custom XML is identified only by Relationship
        self.custom_xml = self._get_parts([],
                                         ['relationships/customXml'])

        self.embedded_controls = self._get_parts(
            ['application/vnd.ms-office.activeX+xml'],
            ['relationships/control'])

        # Note that embedded objects is identified only by Relationship
        self.embedded_objects = self._get_parts([],
                                                ['relationships/oleObject'])

        # Note that embedded packages is identified only by Relationship
        self.embedded_packages = self._get_parts([],
                                                 ['relationships/package'])

        # Identify and provide access to digital signature parts
        self.digital_signatures = self._get_parts(
            ['application/vnd.openxmlformats-package.digital-signaturecertificate,',
             'application/vnd.openxmlformats-package.digital-signature-origin',
             'application/vnd.openxmlformats-package.digital-signaturexmlsignature+xml'],
            ['relationships/digitalsignature/signature',
             'relationships/digitalsignature/certificate',
             'relationships/digitalsignature/origin'])

    def _get_parts(self, content_types, rels):
        """
        Take content_type and relationships as parameters,
        and return all Parts that match.
        """
        parts1 = []
        for ct in content_types:
            for part in self.doc.parts_by_content_type_regex(ct):
                parts1.append(part)
        parts2 = []
        for rel in rels:
            for part in self.doc.parts_by_relationship_type(rel):
                parts2.append(part)
        # Use set() to eliminate duplicates
        return list(set(parts1 + parts2))

    def __repr__(self):
        return "Features of: %s" % self.doc
