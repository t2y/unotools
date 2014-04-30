# -*- coding: utf-8 -*-
from setuptools import setup


VERSION = '0.3.0'

try:
    import pypandoc  
    LONG_DESCRIPTION = '\n'.join([
        pypandoc.convert('README.md', 'rst'),
        pypandoc.convert('CHANGELOG.md', 'rst'),
    ])
except (IOError, ImportError):
    LONG_DESCRIPTION = ''

CLASSIFIERS = [
    'Development Status :: 4 - Beta',
    'Intended Audience :: Developers',
    'Operating System :: MacOS :: MacOS X',
    'Operating System :: POSIX',
    'Programming Language :: Python :: 3',
    'Topic :: Software Development',
    'Topic :: Office/Business :: Office Suites',
]

setup(
    name='unotools',
    version=VERSION,
    description='Interacting with OpenOffice.org/LibreOffice using UNO',
    long_description=LONG_DESCRIPTION,
    classifiers=CLASSIFIERS,
    keywords=['uno', 'pyuno', 'office', 'OOo'],
    author='Tetsuya Morimoto',
    author_email='tetsuya dot morimoto at gmail dot com',
    url='http://bitbucket.org/t2y/unotools',
    license='Apache License 2.0',
    platforms=['unix', 'linux', 'osx'],
    packages=['unotools'],
    #install_requires=['uno', 'pyuno'],
    tests_require=['tox', 'pytest', 'pytest-pep8'],
)