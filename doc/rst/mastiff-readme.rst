OfficeDissector MASTIFF Plugin README 
=====================================

Install MASTIFF 
--------------------

Download MASTIFF source: http://sourceforge.net/projects/mastiff/

Unpack.

::

    $ tar -xzf mastiff-0.6.0.tar.gz

Install prerequisites for MASTIFF: 

::

    $ sudo aptitude install python-setuptools 
    $ sudo aptitude install python-magic

Install (for development testing): 

::

    $ sudo make dev

(For testing OfficeDissector plugins, installing all other plugin
requirements is not necessary. However, errors occur if they are present
without the required modules. Therefore, if not installing all plugin
requirements, remove all Generic plugins from: plugins/GEN/.)

To uninstall: 

::

    $ sudo make dev-clean

Configuring MASTIFF
-----------------------

The mastiff.conf file controls all settings and paths for MASTIFF.
MASTIFF (apparently) searches for it in /etc/mastiff/mastiff.conf,
~/.mastiff.conf, and ./mastiff.conf. You can also force MASTIFF to use a
specific path to mastiff.conf through the -c command line option.

Install OfficeDissector Plugins
-------------------------------

Copy plugins to the MASTIFF source file directory (created when
unpacking the source tarball): 

::

  $ cp path/to/officedissector/mastiff-plugins/* path/to/mastiff/plugins/Office

(MASTIFF's plugins directory is specified in the mastiff.conf file.)

MASTIFF uses the magic file to detect OOXML documents. If running: 

:: 

    $ file test/unit_test/testdocs/test.docx

does not show Microsoft Word 2007+, upgrade the magic library to
identify OOXML files: 

:: 

    $ cp /etc/magic /tmp/magic.backup 
    $ sudo sh -c 'cat mastiff-plugins/magic-ooxml >> /etc/magic'

For more information, see:
http://serverfault.com/questions/338087/making-libmagic-file-detect-docx-files

Alternatively, installing TrID ( http://mark0.net/soft-trid-e.html ) and
editing mastiff.conf to provide the path to TrID should fix this
problem.

Run MASTIFF
-----------

::

    $ mas.py ooxml_file(s) ...

Output can be found in the /path/to/mastiff-source-files/work/log
directory.

If MASTIFF is not being run from the MASTIFF source directory, the user
must specify the location of mastiff.conf using the -c flag. Eg. 

::

    $ mas.py -c /path/to/mastiff/mastiff.conf ooxml_files(s) ...

Also note that if mastiff.conf is using the default settings, MASTIFF
will run only if the current working directory is the MASTIFF source
directory. Otherwise, specify location of the plugins directory in
mastiff.conf.

Plugin Architecture
-------------------

The architecture to create OfficeDissector plugins is located at:
mastiff-plugins/Office-officedissector-skel.py See the beginning of that
file for instructions for creating OfficeDissector plugins.

Example Plugins
---------------

There are three sample plugins:

1. Office-officedissector-multimedia.py
---------------------------------------

To see this plugin run, use: 

::

    $ mas.py -p 'OfficeDissector Extract Multimedia' /path/to/officedissector/test/unit_test/testdocs/037027.pptx

MASTIFF stores results in a folder with path work/log/[md5sum of file]/.

To find the multimedia parts, use: 

::

    $ ls work/log/8aeb72b3751238a37aff319585434327/parts

2. Office-officedissector-urls.py
---------------------------------

To see this plugin run, use: 

:: 

    $ mas.py -p 'OfficeDissector Extract URLs' /path/to/officedissector/test/unit_test/testdocs/037027.pptx

To find the URLs, use: 

::

    $ cat work/log/8aeb72b3751238a37aff319585434327/urls.txt

3. Office-officedissector-embedded-code.py
------------------------------------------

To see this plugin run, use: 

::

    $ mas.py -p 'OfficeDissector Extract Embedded Code' /path/to/officedissector/test/unit_test/testdocs/macros.xlsm

To find the embedded code parts, use: 

::

    $ ls work/log/39c7ca586fefb547b8d7474130ec0fe5/parts

Plugin Tests
------------

To test, first install the plugins as mentioned above. Then:

# Functional tests 

::

    $ cd test/unit_test $ python test_plugins.py PATH_TO_MASTIFF_SOURCE_DIR

# Performance tests 

::

    $ cd test $ python test_plugin_performance.py PATH_TO_MASTIFF_SOURCE_DIR

Note that the tests assume that MASTIFF will output results in
PATH\_TO\_MASTIFF\_SOURCE\_DIR. If you have configured MASTIFF to output
results to the current directory, you may need to specify that instead:

::

    $ cd test/unit_test $ python test_plugins.py . # If MASTIFF is configured to output results to the current directory

For accurate performance results:

1. Disable Zip feedback in mastiff.conf: [ZipExtract] feedback = off.
2. Remove all other plugins outside the MASTIFF source directory,
   especially the ZIP plugins.
3. Remove the Office-officedissector-skel plugin from plugins/Office.

