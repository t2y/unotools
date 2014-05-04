# -*- coding: utf-8 -*-
from pprint import pprint

from com.sun.star.lang import XComponent
from com.sun.star.uno import XComponentContext

from unotools.datatypes import Sequence
from unotools.utils import convert_lowercase_to_camecase


class Component:

    def __init__(self, context: XComponentContext,
                 target_frame_name: str='_blank',
                 search_flags: int=0,
                 arguments: tuple=()):
        self.context = context
        loader = context.load_component_from_url
        self.raw = loader(self.URL, target_frame_name, search_flags, arguments)

    def __getattr__(self, name: str) -> object:
        if self.raw is not None:
            return getattr(self.raw, convert_lowercase_to_camecase(name))
        raise AttributeError

    def _show_attributes(self):
        pprint(dir(self.raw))

    def as_raw(self) -> XComponent:
        return self.raw

    def get_by_index(self, obj: object, index: int) -> object:
        return obj.getByIndex(index)

    def get_by_name(self, obj: object, name: str) -> object:
        return obj.getByName(name)

    def get_count(self, obj: object) -> int:
        return obj.getCount()

    def store_as_url(self, url: str, *values):
        self.raw.storeAsURL(url, self._get_property_values(*values))

    def store_to_url(self, url: str, *values):
        self.raw.storeToURL(url, self._get_property_values(*values))

    def _get_property_values(self, *values) -> Sequence:
        if len(values) == 1 and values[0] is None:
            return Sequence()
        else:
            return Sequence(self.context.make_property_value(*values))
