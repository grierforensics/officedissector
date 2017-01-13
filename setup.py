#!/usr/bin/env python

from setuptools import setup

setup(
    author='Grier Forensics',
    author_email='jdgrier@grierforensics.com',
    description="""OfficeDissector is a parser library for static security analysis of OOXML documents.""",
    name='officedissector',
    version='1.0',
    test_suite="test",
    url='http://www.officedissector.com/',
    packages=['officedissector'],
    install_requires=['lxml']
)
