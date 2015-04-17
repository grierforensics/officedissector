Installing OfficeDissector
-------------------------------------

OfficeDissector requires the lxml package and Python version 2.7. OfficeDissector can be installed using pip (recommended) or python
setup:

::

    $ sudo pip install /path/to/thisfolder # Recommended, as pip supports uninstall
    $ sudo python setup.py install # Alternative

Alternatively, to use OfficeDissector without installing it, set the ``PYTHONPATH`` to
the ``officedissector`` directory:

::

    $ export PYTHONPATH=/path/to/thisfolder


Documentation
-------------

To view OfficeDissector documentation, open in a browser:

::

    $ doc/html/index.html

Testing
-------

To test, first set PYTHONPATH or install ``officedissector`` as
described above. Then:

::

    # Unit tests
    $ cd test/unit_test
    $ python test_officedissector.py

    # Smoke tests
    $ cd test
    $ python smoke_tests.py

The smoke tests will create log files with more information about them.

