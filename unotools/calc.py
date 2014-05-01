# -*- coding: utf-8 -*-
from unotools.datatypes import Sequence


class Calc:

    URL = 'private:factory/scalc'

    def __init__(self, context, target_frame_name='_blank', search_flags=0,
                 arguments=()):
        self.context = context
        self.component = context.load_component_from_url(Calc.URL,
                            target_frame_name, search_flags, arguments)

    def close(self):
        self.component.close(True)

    @property
    def sheets(self):
        return self.component.getSheets()

    def get_sheets_by_index(self, index):
        return self.sheets.getByIndex(index)

    def get_sheets_by_name(self, name):
        return self.sheets.getByName(name)

    def get_sheets_count(self):
        return self.sheets.getCount()

    def get_number_formats(self):
        return self.component.getNumberFormats()

    def store_as_url(self, url, *values):
        self.component.storeAsURL(url, self._get_property_values(*values))

    def store_to_url(self, url, *values):
        self.component.storeToURL(url, self._get_property_values(*values))

    def _get_property_values(self, *values):
        if len(values) == 1 and values[0] is None:
            return Sequence()
        else:
            return Sequence(self.context.make_property_value(*values))
