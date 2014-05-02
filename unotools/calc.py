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
    def set_rows(self, sheet, x, y, data, method_names):
        for i, datum in enumerate(data):
            self.set_rows_cell_data(sheet, x, y, [datum], method_names[i])
            x += 1

    def set_columns(self, sheet, x, y, data, method_names):
        for i, datum in enumerate(data):
            self.set_rows_cell_data(sheet, x, y, [datum], method_names[i])
            y += 1

    def set_rows_cell_data(self, sheet, x, y, data, method_name):
        for datum in data:
            methodcaller(method_name, datum)(sheet.getCellByPosition(x, y))
            y += 1

    def set_columns_cell_data(self, sheet, x, y, data, method_name):
        for datum in data:
            methodcaller(method_name, datum)(sheet.getCellByPosition(x, y))
            x += 1

    def set_rows_str(self, sheet, x, y, data):
        self.set_rows_cell_data(sheet, x, y, data, 'setString')

    def set_columns_str(self, sheet, x, y, data):
        self.set_columns_cell_data(sheet, x, y, data, 'setString')

    def set_rows_value(self, sheet, x, y, data):
        self.set_rows_cell_data(sheet, x, y, data, 'setValue')

    def set_columns_value(self, sheet, x, y, data):
        self.set_columns_cell_data(sheet, x, y, data, 'setValue')

    def set_rows_formula(self, sheet, x, y, data):
        self.set_rows_cell_data(sheet, x, y, data, 'setFormula')

    def set_columns_formula(self, sheet, x, y, data):
        self.set_columns_cell_data(sheet, x, y, data, 'setFormula')

    # charts operation
