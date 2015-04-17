#!/usr/bin/env python

"""
An interface to the OOXML Document as a Zip file,
and a lightweight wrapper around Zipfile.
"""

__author__ = 'Brandon Gordon'
__email__ = 'bgordon@grierforensics.com'

import zipfile


class Zip(object):
    """
    An interface to the OOXML Document as a Zip file,
    and a lightweight wrapper around the `ZipFile` module.

    :ivar filepath: Path to the Zip (OOXML) file.

    :ivar zippartsinfo: A list containing a `ZipInfo` object for each
        member of the Zip file. The `ZipInfo` object contains all
        information about the member of the Zip file.

    :ivar comment: The comment text associated with the Zip file.
    """

    def __init__(self, filepath):
        """
        Initialize zip attributes.

        :param filepath: the path to the OOXML Document
        :type filepath: string
        """
        self.filepath = filepath
        self._zipobj = zipfile.ZipFile(self.filepath, 'r')

        self.zippartsinfo = self._zipobj.infolist()

        self.comment = self._zipobj.comment

    def testzip(self):
        """
        Test zip CRC value.

        :raises ZipCRCError: If the Zip CRC is incorrect
        """
        if self._zipobj.testzip():
            raise ZipCRCError("Zip file CRC is invalid")

    def namelist(self):
        """
        Get list of files in Zip archive.

        :return: list of names of files in the Zip archive
        """
        return self._zipobj.namelist()

    def part_extract(self, partname):
        """
        Extract part from the Zip archive.

        :param partname: name of the :class:`~officedissector.part.Part` (member of the Zip archive)
            to extract
        :type partname: string
        :return: file-like object of the member of the Zip archive.
        """
        return self._zipobj.open(partname.lstrip('/'))

    def part_info(self, partname):
        """
        Get `ZipInfo` object for :class:`~officedissector.part.Part`.

        :param partname: name of :class:`~officedissector.part.Part` (member of the Zip archive)
        :type partname: string
        :return: `ZipInfo`
        """
        # Members of Zip archive do not have leading '/'
        return self._zipobj.getinfo(partname.lstrip('/'))

    def __repr__(self):
        return "Zip File: %s" % self.filepath


class ZipCRCError(Exception):
    """Raise an Exception when Zip CRC value is invalid."""

    def __init__(self, msg):
        Exception.__init__(self, msg)
        self.msg = msg

    def __str__(self):
        return repr(self.msg)