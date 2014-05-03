# -*- coding: utf-8 -*-
from com.sun.star.lang import XComponent
from com.sun.star.uno import XComponentContext
from com.sun.star.uno import XInterface

from unotools.datatypes import Sequence


class Component:

    def __init__(self, context: XComponentContext,
                 target_frame_name: str='_blank',
                 search_flags: int=0,
                 arguments: tuple=()):
        self.context = context
        loader = context.load_component_from_url
        self.raw = loader(self.URL, target_frame_name, search_flags, arguments)

    def close(self):
        self.raw.close(True)

    def as_raw(self) -> XComponent:
        return self.raw

    def get_string(self) -> str:
        return self.raw.getString()

    def set_string(self, text: str):
        self.raw.setString(text)

    def get_by_index(self, obj: object, index: int) -> object:
        return obj.getByIndex(index)

    def get_by_name(self, obj: object, name: str) -> object:
        return obj.getByName(name)

    def get_count(self, obj: object) -> int:
        return obj.getCount()

    def get_title(self):
        return self.raw.getTitle()

    def create_instance(self, name: str) -> XInterface:
        return self.raw.createInstance(name)

    def store_as_url(self, url: str, *values):
        self.raw.storeAsURL(url, self._get_property_values(*values))

    def store_to_url(self, url: str, *values):
        self.raw.storeToURL(url, self._get_property_values(*values))

    def _get_property_values(self, *values) -> Sequence:
        if len(values) == 1 and values[0] is None:
            return Sequence()
        else:
            return Sequence(self.context.make_property_value(*values))
