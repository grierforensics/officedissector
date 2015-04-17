#!/usr/bin/env python
"""
OfficeDissector-Multimedia Plug-in

Extracts all multimedia from Document: images, sounds,
videos, and fonts.

Outputs multimedia Parts into the parts directory in the
MASTIFF log directory (specified in mastiff.conf; default is
work/log/[md5 of filename]/parts).

Parts have '+' substituted for '/'. Eg. a part '/word/document.xml'
becomes 'word+document.xml'.

"""

__author__ = "Brandon Gordon"
__email__ = "bgordon@grierforensics.com"

import logging
import os

from zipfile import BadZipfile

import mastiff.category.office as office
import officedissector

class OfficeDissectorSkel(office.OfficeCat):
    """OfficeDissector-Multimedia Plug-in"""

    def __init__(self):
        """Initialize the plugin."""
        office.OfficeCat.__init__(self)

    def analyze(self, config, filename):
        """Analyze the file."""

        # make sure we are activated
        if not self.is_activated:
            return False
        log = logging.getLogger('Mastiff.Plugins.' + self.name)
        log.info('Starting execution on %s.' % filename)

        try:
            doc = officedissector.Document(filename)
        except BadZipfile:
            # In mastiff.conf, if [ZipExtract] feedback = on,
            # the plugins are run again on unzipped Parts. If a Part
            # looks like an Office document, but isn't OOXML, (eg. an
            # embedded object) a BadZipFile exception is raised.
            log.error('%s is not a ZIP file.' % filename)
            return False
        except Exception, err:
            log.error('Error opening document: %s', err)
            return False


        # create the dir if it doesn't exist
        out_part_dir = config.get_var('Dir', 'log_dir') + os.sep + 'parts'
        if not os.path.exists(out_part_dir):
            try:
                os.makedirs(out_part_dir)
            except IOError,  err:
                log.error('Unable to create dir %s: %s' % (out_part_dir, err))
                return False

        multimedia = (doc.features.images +
                      doc.features.sounds +
                      doc.features.videos +
                      doc.features.fonts)
        for part in multimedia:
            self.output_part(out_part_dir, part)

        log.debug('Successfully ran %s.', self.name)

        return True

    def output_file(self, outdir, data):
        """Print output from analysis to a file."""
        return True

    def output_part(self, outdir, part):
        """Write the contents of any Part to a file."""
        log = logging.getLogger('Mastiff.Plugins.' + self.name)

        part_stream = part.stream().read()

        try:
            part_file = open(
                os.path.join(outdir, part.name.strip('/').replace('/', '+')),
                'w')

        except IOError, err:
            log.error('Write error: %s' % err)
            return False

        part_file.write(part_stream)

        part_file.close()
        return True


