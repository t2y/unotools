# -*- coding: utf-8 -*-
from unotools.datatypes import Sequence


class Component:

    def __init__(self, context, target_frame_name='_blank', search_flags=0,
                 arguments=()):
        self.context = context
        loader = context.load_component_from_url
        self.component = loader(self.URL, target_frame_name, search_flags,
                                arguments)

    def close(self):
        self.component.close(True)

    def store_as_url(self, url, *values):
        self.component.storeAsURL(url, self._get_property_values(*values))

    def store_to_url(self, url, *values):
        self.component.storeToURL(url, self._get_property_values(*values))

    def _get_property_values(self, *values):
        if len(values) == 1 and values[0] is None:
            return Sequence()
        else:
            return Sequence(self.context.make_property_value(*values))
