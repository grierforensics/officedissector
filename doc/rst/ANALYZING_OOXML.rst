Analyzing OOXML with OfficeDissector
====================================

OOXML files can be thought of at six levels of abstraction:

-  As a file
-  As a Zip archive
-  As a set of Parts (a Part has a *name* and an array of data bytes
   called a *stream*)
-  As a set of Parts with data in well defined formats (usually XML)
-  As a graph of Parts, connected by Relationships, and having
   associated metadata (specifically Content-Type)
-  As a document with Features and CoreProperties

Each abstraction level completely contains all the levels beneath it;
there are no leaks. That is, all Features are implemented as
Parts+Relationships, all Parts are contained in the Zip file, etc. This
concept will be explained more below.

Different types of analysis can be done at each of these levels.

File
----

At its simplest, an OOXML file is a single file; that is, a filename and
a sequence of bytes. OfficeDissector's :class:`~officedissector.doc.Document` class takes this file
and exposes its deeper content:

::

    import officedissector
    doc = officedissector.doc.Document('test/fraunhoferlibrary/Artikel.docx') # Returns a Document object

Note that the filename's extension is signficant and affects the
behavior of Microsoft Office. Member variables :class:`doc.type <officedissector.doc.Document>`,
:class:`doc.is_macro_enabled <officedissector.doc.Document>` and :class:`doc.is_template <officedissector.doc.Document>` expose the meaning of
the filename extension.

Zip file
--------

An OOXML file is a Zip archive (each members of this archive is called a
Part, described below). Method :func:`officedissector.doc.Document.zip` provides a :class:`~officedissector.zip.Zip` object
which provides access at this level of abstraction. However, most
analysis is best performed at the deeper levels of abstraction below.

Part
----

:class:`Parts <officedissector.part.Part>` are the heart of OOXML. **The entire Document is defined by its
Parts** (other than the limited information provided by the filename's
extension and Zip metadata described above). Even the Parts' own
metadata is stored in Parts (as will be described below). Thus,
**analyzing OOXML is about analyzing Parts**.

:class:`Document.parts <officedissector.doc.Document>` provides a List of all Parts in the document, and
:class:`Document.parts_by_name <officedissector.doc.Document>` provides a Dictionary of Parts by their name.

At their simplest level, Parts have only two properties: a *name*
(exposed through :class:`Part.name <officedissector.part.Part>`) and an array of data bytes called a
*stream* (exposed through :func:`officedissector.part.Part.stream`).

::

    for p in doc.parts:
        print p.name # p.name returns a String
        print p.stream().read(10) # p.stream() returns a File-like object

Note that, with a few exceptions, the Part's name is irrelevant and will
not affect the behavior of Office.
Parts roles are instead determined by their Content-Type and
Relationships (described below).

Content-Type
------------

All :class:`Parts <officedissector.part.Part>` have a Content-Type, such as ``image/png`` or
``application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml``,
exposed via the :func:`officedissector.part.Part.content_type` method. (Content-Types are
defined by a single Part with a well known name,
``[Content_Types.xml]``, which OfficeDissector automatically parses.)

Relationships
-------------

Parts have :class:`Relationships <officedissector.rel.Relationship>` to other parts (or external resources), forming
a graph-like structure. (Relationships themselves are defined by
dedicated ``.rels`` Parts, which OfficeDissector automatically finds and
parses.)

:class:`officedissector.doc.Document.relationships <officedissector.doc.Document>` provides a List of all Relationships in the
Document. :func:`officedissector.part.Part.relationships_in` provides a List of all
Relationships pointing *to* a Part; and :func:`officedissector.part.Part.relationships_out`
provides a List of all Relationships *from* a Part.

Finding Parts
-------------

OfficeDissector provides several ways to find Parts of interest:

-  :func:`officedissector.doc.Document.main_part` returns the main Part; each Document has
   exactly one main Part
-  :func:`officedissector.doc.Document.parts_by_content_type` and
   :func:`officedissector.doc.Document.parts_by_content_type_regex` allowing finding all Parts
   with particular Content-Type
-  :func:`officedissector.doc.Document.parts_by_relationship_type` allows finding all Parts
   which have an incoming Relationship of a particular type.

In nearly every case, Office's treatment of a Part is soley determined
by the Part's Content-Type and incoming Relationship types. Thus, these
methods provide the best way to perform most analysis.

XML Parts
---------

Most Parts are in XML format. OfficeDissector will parse the XML for
you, using the :func:`officedissector.part.Part.xml` method, or perform XPath queries
(recommended), using the :func:`officedissector.part.Part.xpath` method.
XPath is a very powerful tool which makes most OOXML analysis much
simpler.

OOXML makes heavy use of XML Namespaces. XML Namespaces can be confusing
if you do not have experience with them. See :doc:`XMLNamespaces.pdf <namespaces>` for a
quick introduction.

Other Parts
-----------

Multimedia and embedded objects are typically stored in their own Parts,
one Part per object. For example, each image will comprises its own
Part, with the appropriate Content-Type (e.g. ``image/png``) and
Relationships. Since each image is kept separate, in its native format,
and with its Content-Type exposed, analyzing them is easy: The image's
data is readable as :func:`officedissector.part.Part.stream`. For example:

::

    jpegs = doc.parts_by_content_type('image/jpeg')
    jpegs[0].name # Returns '/word/media/image1.jpeg'
    jpegs[0].stream().read(10) # Returns '\xff\xd8\xff\xe0\x00\x10JFIF', the first 10 bytes of the JPEG data

Features and CoreProperties
---------------------------

The :class:`~officedissector.features.Features` object (retrieved via :class:`Document.features <officedissector.doc.Document>`) provides
access to common Document features, such as macros, images, videos, and
sounds. This interface provides convenient access to the features most
relevant to security analysis. (It's important to realize that all of
these features are simply Parts found by their Content-Type and
Relationships.)

The :class:`~officedissector.core_properties.CoreProperties` object (retrieved via :class:`Document.core_properties <officedissector.doc.Document>`)
parses and exposes the Document's Core Properties, such as ``creator``
and ``modified``; these properties are very useful for security analysis
and forensics.

How do I...?
============

To **get started**, install OfficeDissector (see :doc:`Installing <install>`), and begin with these two lines:

::

    import officedissector
    doc = officedissector.doc.Document('path/to/your/ooxml.docx') # Returns a Document object

To **interactively analyze an OOXML document**, use ipython. See :doc:`Usage <usage>`
for an example.

To **automatically analyze a large volume of documents**, use plugins.
See :doc:`mastiff-plugins/README.txt <mastiff-readme>`.

To **learn OfficeDissector**, use ipython, and press the TAB key to see
the methods available. (This is demonstrated in :doc:`Usage <usage>`.) Look at the full :doc:`API docs <api>` 
for more details.

If you want to **retrieve specific features and properties of a
document**, use the `Features and CoreProperties`_ interface.

The best way to **explore the behavior of Office** is to retrieve Parts
by their Content-Type and Relationships; see `Finding Parts`_ above. As far
as we know, all of Office's behavior can be reconstructed through this
technique. (Office does also looks at the filename's extension; see `File`_ above.) 
**This is therefore the most general purpose and powerful mode of analysis**. (This is the method OfficeDissector uses
internally to find Features, such as multimedia.)

To **drill deeper into an XML Part**, use :func:`officedissector.part.Part.xpath`, paying
careful attention to XML namespaces; see `XML Parts`_ above.
(This is the method OfficeDissector uses internally to find and parse
CoreProperties, such as ``creator``.)

To **export the data into other tools**, use the :func:`to_json() <officedissector.doc.Document.to_json>` method,
which nearly all OfficeDissector objects provide.
