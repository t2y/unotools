# -*- coding: utf-8 -*-
import functools
from os.path import realpath

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
