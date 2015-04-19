# OfficeDissector

OfficeDissector is a parser library for static security analysis of Office Open XML (OOXML) Documents,
created by Grier Forensics for the Cyber System Assessments Group at MIT's Lincoln Laboratory.

OfficeDissector is the first parser designed specifically for security analysis of OOXML documents.  It exposes all internals, including 
document properties, parts, content-type, relationships, embedded macros and multimedia, and comments, and more. 
It provides full JSON export, and a MASTIFF based plugin architecture.  It also includes a nearly 600 MB test corpus, unit tests with nearly 
100% coverage, smoke tests running against the entire corpus, and simple, well factored, fully commented code 

## Install

OfficeDissector requires Python 2.7 and the lxml package.

The easiest way to install OfficeDissector is to use pip to automatically download and install it:

    $ sudo pip install lxml # If you haven't installed lxml already
    $ sudo pip install officedissector

Alternatively, you can download OfficeDissector from [github](https://github.com/grierforensics/officedissector/) or as a [zip](https://github.com/grierforensics/officedissector/archive/master.zip), and install your local copy, using either pip (recommended) or python setup:

    $ sudo pip install /path/to/thisfolder # Recommended, as pip supports uninstall
    $ sudo python setup.py install # Alternative

Finally, to use OfficeDissector without installing it, download it and set the `PYTHONPATH` to the `officedissector` directory:

    $ export PYTHONPATH=/path/to/thisfolder

## Documentation

To view OfficeDissector documentation, open in a browser:

    $ doc/html/index.html

## Testing

To test, first set PYTHONPATH or install `officedissector` as described above.  Then:

    # Unit tests
    $ cd test/unit_test
    $ python test_officedissector.py

    # Smoke tests
    $ cd test
    $ python smoke_tests.py

The smoke tests will create log files with more information about them.

## MASTIFF Plugins

To find more information about the MASTIFF architecture and sample plugins, see
`mastiff-plugins/README.txt`.

## Usage

Below is an ipython session demonstrating usage of OfficeDissector:

    $ ipython
    In [1]: import officedissector
    In [2]: doc = officedissector.doc.Document('test/fraunhoferlibrary/Artikel.docx')
    In [4]: doc.is_macro_enabled
    Out[4]: False

    In [5]: doc.is_template
    Out[5]: False

    In [6]: mp = doc.main_part()
    In [7]: mp.content_type()
    Out[7]: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'

    In [9]: mp.name
    Out[9]: '/word/document.xml'

    In [10]: mp.content_type()
    Out[10]: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'

    # We can read the part's stream of data:
    In [17]: mp.stream().read(200)
    Out[17]: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-c'

    # Or use XPath to parse it:
    In [33]: t = mp.xpath('//w:t', {'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main"})
    In [37]: t[2].text
    Out[37]: u'Das vorliegende Dokument ist ein Beispiel f\xfcr einen zur Publikation in einer Zeitschrift vorgesehenen Artikel. Es verwendet f\xfcr Autor und Titel in den Dokumenteigenschaften festgelegte Eintr\xe4ge.'

    # All Relationships in and out are exposed:
    In [38]: mp.relationships_in()
    Out[38]: [Relationship [rId1] (source Part [RootPart])]

    In [39]: mp.relationships_out()
    Out[39]:
    [Relationship [rId8] (source Part [/word/document.xml]),
     Relationship [rId13] (source Part [/word/document.xml]),
     Relationship [rId3] (source Part [/word/document.xml]),
     ...
     Relationship [rId14] (source Part [/word/document.xml])]

    In [40]: rel = mp.relationships_out()[0]
    In [43]: rel.type
    Out[43]: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes'

    In [46]: endnotes = rel.target_part
    In [48]: endnotes.content_type()
    Out[48]: 'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml'

    # Any Part (or the entire Document) can be exported to JSON:
    In [50]: print endnotes.to_json()
    {
        "content-type": "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
        "uri": "/word/endnotes.xml",
        "relationships_out": [],
        "relationships_in": [
            "Relationship [rId8] (source Part [/word/document.xml])"
        ]
    }

    # Features are automatically exposed:
    In [55]: doc.features.[TAB]
    ...
    doc.features.comments
    doc.features.custom_properties
    doc.features.custom_xml
    doc.features.digital_signatures
    doc.features.doc
    doc.features.embedded_controls
    doc.features.embedded_objects
    doc.features.embedded_packages
    doc.features.fonts
    doc.features.get_parts
    doc.features.get_union
    doc.features.images
    doc.features.macros
    doc.features.sounds
    doc.features.videos

    In [55]: doc.features.images
    Out[55]: [Part [/word/media/image1.jpeg]]

    In [56]: image = doc.features.images[0]
    In [58]: image.content_type()
    Out[58]: 'image/jpeg'

    # We can export the binary data to JSON as well, by setting include_stream = True:
    In [61]: print image.to_json(include_stream = True)
    {
        "stream_b64": "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAFAAUDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD3uGGO3iWKJdqL0Gc0UUUAf//Z",
        "content-type": "image/jpeg",
        "uri": "/word/media/image1.jpeg",
        "relationships_out": [],
        "relationships_in": [
            "Relationship [rId1] (source Part [/word/theme/theme1.xml])"
        ]
    }

    # Check for macros:
    In [62]: doc.features.macros
    Out[62]: []

    # Or comments:
    In [63]: doc.features.comments
    Out[63]: []

    # Core properties are exposed:
    In [64]: doc.core_properties.[TAB]
    ...
    doc.core_properties.content_status
    doc.core_properties.core_prop_part
    doc.core_properties.created
    doc.core_properties.creator
    doc.core_properties.description
    doc.core_properties.identifier
    doc.core_properties.keywords
    doc.core_properties.language
    doc.core_properties.last_modified_by
    doc.core_properties.last_printed
    doc.core_properties.modified
    doc.core_properties.name
    doc.core_properties.parse_all
    doc.core_properties.parse_prop
    doc.core_properties.revision
    doc.core_properties.subject
    doc.core_properties.title
    doc.core_properties.version
    doc.core_properties.category

    In [68]: doc.core_properties.modified
    Out[68]: '2009-12-04T14:47:00Z'

## Analyzing OOXML

See `doc/txt/ANALYZING_OOXML.txt` for a quick start guide on how to use 
OfficeDissector to analyze OOXML documents.

## API

For more details about OfficeDissector, see the API - `doc/html/rst/api.html` documentation.

## More Information

See http://www.officedissector.com for more information on the project.
