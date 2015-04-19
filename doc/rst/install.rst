Installing OfficeDissector
-------------------------------------

OfficeDissector requires Python 2.7 and the lxml package.

The easiest way to install OfficeDissector is to use pip to automatically download and install it:

::

    $ sudo pip install lxml # If you haven't installed lxml already
    $ sudo pip install officedissector

Alternatively, you can download OfficeDissector from `github <https://github.com/grierforensics/officedissector>`_. or as a `zip <https://github.com/grierforensics/officedissector/archive/master.zip>`_ and install your local copy, using either pip (recommended) or python setup:

::

    $ sudo pip install /path/to/thisfolder # Recommended, as pip supports uninstall
    $ sudo python setup.py install # Alternative

Finally, to use OfficeDissector without installing it, download it and set the ``PYTHONPATH`` to the ``officedissector`` directory:

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

