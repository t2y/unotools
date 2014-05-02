# -*- coding: utf-8 -*-
import csv
from os.path import basename


class CsvFile:

    def __init__(self, path, mode='r', encoding='utf-8', has_header=False,
                 convert_func=lambda x: x):
        self.path = path
        self.file_name = basename(path)
        self.mode = mode
        self.encoding = encoding
        self.header = None
        self.has_header = has_header
        self.convert_func = convert_func

    def read(self):
        with open(self.path, mode=self.mode, encoding=self.encoding) as f:
            reader = csv.reader(f)
            if self.has_header:
                self.header = next(reader)
            for data in reader:
                yield list(map(self.convert_func, data))