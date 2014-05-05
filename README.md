# unotools

Python libraries for interacting with OpenOffice.org/LibreOffice using the "UNO bridge".


## How to install

### On Ubuntu 14.04 

Install libreoffice, uno library and python3:

    $ sudo aptitude install -y libreoffice libreoffice-script-provider-python uno-libs3 python3-uno python3

I like virtualenvwrapper to make temporary environment:

    $ sudo aptitude install -y virtualenvwrapper
    $ mkvirtualenv -p /usr/bin/python3.4 --system-site-packages tmp3

Confirm importing uno module:

    (tmp3)$ python 
    Python 3.4.0 (default, Apr 11 2014, 13:05:11) 
    [GCC 4.8.2] on linux
    Type "help", "copyright", "credits" or "license" for more information.
    >>> import uno

Install unotools from PyPI:

    (tmp3)$ pip install unotools

## How to use

Startup libreoffice:

    (tmp3)$ soffice --accept='socket,host=localhost,port=8100;urp;StarOffice.Service'

Download sample-scripts from https://bitbucket.org/t2y/unotools/raw/default/sample-scripts:

    (tmp3)$ python sample-scripts/writer-sample1.py -s localhost
    (tmp3)$ python sample-scripts/calc-sample1.py -s localhost -d sample-scripts/datadir/
    (tmp3)$ ls
    sample-calc.html  sample-calc.pdf  sample-calc_html_eaf26d01.png
    sample-scripts  sample-writer.html  sample-writer.pdf
    sample.csv  sample.doc  sample.xls

