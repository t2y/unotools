# -*- coding: utf-8 -*-
import functools


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
    if len(path) > 1:
        if path[1:2] == ':':
            path = '/' + path[0] + '|' + path[2:]
    return 'file://' + path.replace('\\', '/')
