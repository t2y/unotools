# -*- coding: utf-8 -*-
from operator import methodcaller

from com.sun.star.awt import Rectangle
from com.sun.star.chart import XChartDocument
from com.sun.star.chart import XDiagram
from com.sun.star.sheet import XCellRangeAddressable
from com.sun.star.sheet import XSpreadsheet
from com.sun.star.sheet import XSpreadsheets
from com.sun.star.table import CellRangeAddress
from com.sun.star.table import XCellRange
from com.sun.star.table import XTableChart
from com.sun.star.table import XTableCharts
from com.sun.star.util import XNumberFormats

from unotools.component import Component
from unotools.datatypes import Sequence


class ChartDocument(Component):

    def __init__(self, chart_document: XChartDocument):
        self.raw = chart_document

    def set_diagram(self, diagram: XDiagram):
        self.raw.setDiagram(diagram)

    def get_diagram(self) -> XDiagram:
        return self.raw.getDiagram()


class TableChart(Component):

    def __init__(self, chart: XTableChart):
        self.raw = chart

    def get_embedded_object(self) -> ChartDocument:
        return ChartDocument(self.raw.getEmbeddedObject())


class Spreadsheet(Component):

    # sheet operation
    def __init__(self, sheet: XSpreadsheet):
        self.raw = sheet

    def set_rows_cell_data(self, x: int, y: int, data: list, method_name: str):
        for datum in data:
            methodcaller(method_name, datum)(self.raw.getCellByPosition(x, y))
            y += 1

    def set_columns_cell_data(self, x: int, y: int, data: list,
                              method_name: str):
        for datum in data:
            methodcaller(method_name, datum)(self.raw.getCellByPosition(x, y))
            x += 1

    def set_rows(self, x: int, y: int, data: list, method_names: str):
        for i, datum in enumerate(data):
            self.set_rows_cell_data(x, y, [datum], method_names[i])
            x += 1

    def set_columns(self, x: int, y: int, data: list, method_names: str):
        for i, datum in enumerate(data):
            self.set_rows_cell_data(x, y, [datum], method_names[i])
            y += 1

    def set_rows_str(self, x: int, y: int, data: list):
        self.set_rows_cell_data(x, y, data, 'setString')

    def set_columns_str(self, x: int, y: int, data: list):
        self.set_columns_cell_data(x, y, data, 'setString')

    def set_rows_value(self, x: int, y: int, data: list):
        self.set_rows_cell_data(x, y, data, 'setValue')

    def set_columns_value(self, x: int, y: int, data: list):
        self.set_columns_cell_data(x, y, data, 'setValue')

    def set_rows_formula(self, x: int, y: int, data: list):
        self.set_rows_cell_data(x, y, data, 'setFormula')

    def set_columns_formula(self, x: int, y: int, data: list):
        self.set_columns_cell_data(x, y, data, 'setFormula')

    def get_cell_range_by_name(self, range_: str) -> XCellRange:
        return self.raw.getCellRangeByName(range_)

    def get_cell_range_by_position(self, left: int, top: int, right: int,
                                   bottom: int) -> XCellRange:
        return self.raw.getCellRangeByPosition(left, top, right, bottom)

    def get_range_address(self, cell_range: XCellRange
                          ) -> XCellRangeAddressable:
        return cell_range.getRangeAddress()

    # charts operation
    @property
    def charts(self) -> XTableCharts:
        return self.raw.getCharts()

    def get_charts_count(self) -> int:
        return self.get_count(self.charts)

    def add_charts_new_by_name(self, name: str,
                               rect: Rectangle, ranges: CellRangeAddress,
                               column_headers: bool, row_headers: bool):
        self.charts.addNewByName(name, rect, Sequence(ranges),
                                 column_headers, row_headers)

    def get_chart_by_index(self, index: int) -> TableChart:
        return TableChart(self.get_by_index(self.charts, index))

    def get_chart_by_name(self, name: str) -> TableChart:
        return TableChart(self.get_by_name(self.charts, name))


class Calc(Component):

    URL = 'private:factory/scalc'

    # sheets operation
    @property
    def sheets(self) -> XSpreadsheets:
        return self.raw.getSheets()

    def get_sheets_count(self) -> int:
        return self.get_count(self.sheets)

    def get_sheet_by_index(self, index: int) -> Spreadsheet:
        return Spreadsheet(self.get_by_index(self.sheets, index))

    def get_sheet_by_name(self, name: str) -> Spreadsheet:
        return Spreadsheet(self.get_by_name(self.sheets, name))

    def get_number_formats(self) -> XNumberFormats:
        return self.raw.getNumberFormats()

    def insert_sheets_new_by_name(self, name: str, position: int):
        self.sheets.insertNewByName(name, position)

    def insert_multisheets_new_by_name(self, data: list, position: int):
        for i, datum in enumerate(data):
            self.insert_sheets_new_by_name(datum, position + i)

    def remove_sheets_by_name(self, name: str):
        self.sheets.removeByName(name)
