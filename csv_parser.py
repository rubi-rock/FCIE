import csv

import csv
from io import StringIO

from constants import CSVBlocks, ExcelBlock, MatchingBlocks

class CSVParser(object):
    def __init__(self, filename):
        self.__values = {}
        self.__parse(filename)
        pass

    def __parse(self, filename):
        excel_block = None
        read_headers = False
        with open(filename, mode='rt', encoding='iso-8859-1') as csv_file:
            for line in iter(csv_file):
                # block detection
                if line.startswith("==="):
                    line = line.replace('=', '').replace('\n', '')
                    # process only known blocks
                    if line in MatchingBlocks.keys():
                        read_headers = True
                        excel_block = MatchingBlocks[line]
                        self.__values[excel_block] = []
                        continue
                    else:
                        excel_block = None

                # skip headers
                if read_headers:
                    read_headers = False
                    continue

                # add values to block
                if excel_block is not None:
                    csv_stream = StringIO(line)
                    csv_line = csv.reader(csv_stream, delimiter=',', quotechar='"')
                    self.__values[excel_block].append(next(csv_line))

    @property
    def values(self):
        return self.__values



