# -*- coding: utf-8 -*-
import functools
import io
import os
from os.path import join as pathjoin, realpath

import uno
import pyuno


def cached_property(func):
    return property(functools.lru_cache()(func))


def get_annotation_to_kwargs(klass, func_name, values):
    ann = getattr(klass, func_name).__annotations__
    return dict((key, values.get(key)) for key in ann.keys()
                if values.get(key) is not None)


def set_kwargs(obj, kwargs):
    for key, value in kwargs.items():
        if value is not None:
            setattr(obj, key.title(), value)


def get_file_list(path):
    for root, dirs, files in os.walk(path):
        for filename in files:
            yield pathjoin(root, filename)


def read_file(path, mode='r', encoding='utf-8'):
    with io.open(path, mode=mode, encoding=encoding) as f:
        for line in f:
            yield line.strip()


def convert_path_to_url(path):
    """
    >>> convert_path_to_url('/var/tmp/libreoffice')
    'file:///var/tmp/libreoffice'
    """
    return pyuno.systemPathToFileUrl(realpath(path))


def convert_url_to_path(url):
    """
    >>> convert_url_to_path('file:///var/tmp/libreoffice')
    '/var/tmp/libreoffice'
    """
    return pyuno.fileUrlToSystemPath(url)
