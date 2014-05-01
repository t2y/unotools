# -*- coding: utf-8 -*-
from operator import methodcaller

from unotools.component import Component


class Calc(Component):

    URL = 'private:factory/scalc'

    # sheets operation
    @property
    def sheets(self):
        return self.component.getSheets()

    def get_sheets_by_index(self, index: int):
        return self.sheets.getByIndex(index)

    def get_sheets_by_name(self, name: str):
        return self.sheets.getByName(name)

    def get_sheets_count(self):
        return self.sheets.getCount()

    def get_number_formats(self):
        return self.component.getNumberFormats()

    def insert_sheets_new_by_name(self, name: str, position: int):
        self.sheets.insertNewByName(name, position)

    def insert_multisheets_new_by_name(self, data: list, position: int):
        for i, datum in enumerate(data):
            self.insert_sheets_new_by_name(datum, position + i)

    def remove_sheets_by_name(self, name: str):
        self.sheets.removeByName(name)

    # a sheet operation
    def _set_cell_data(self, sheet, x, y, data, method_name, is_rows):
        for datum in data:
            methodcaller(method_name, datum)(sheet.getCellByPosition(x, y))
            if is_rows:
                y += 1
            else:
                x += 1

    def set_rows_str(self, sheet, x, y, data):
        self._set_cell_data(sheet, x, y, data, 'setString', True)

    def set_columns_str(self, sheet, x, y, data):
        self._set_cell_data(sheet, x, y, data, 'setString', False)

    def set_rows_value(self, sheet, x, y, data):
        self._set_cell_data(sheet, x, y, data, 'setValue', True)

    def set_columns_value(self, sheet, x, y, data):
        self._set_cell_data(sheet, x, y, data, 'setValue', False)

    def set_rows_formula(self, sheet, x, y, data):
        self._set_cell_data(sheet, x, y, data, 'setFormula', True)

    def set_columns_formula(self, sheet, x, y, data):
        self._set_cell_data(sheet, x, y, data, 'setFormula', False)

    # charts operation
