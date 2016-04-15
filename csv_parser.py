import csv
from io import StringIO
from constants import ExcelBlock

class CSVParser(object):
    def __init__(self, filename):
        self.__values = {}
        self.__spare = 5
        self.__parse(filename)
        self.__add_spare()
        pass

    def __add_spare(self):
        for list in self.__values.values():
            if len(list) == 0:
                continue
            l = len(list[0])
            for i in range(0, self.__spare):
                list.append([None] * l)

    def __parse(self, filename):
        excel_block = None
        read_headers = False
        with open(filename, mode='rt', encoding='iso-8859-1') as csv_file:
            for line in iter(csv_file):
                # block detection
                if line.startswith("==="):
                    line = line.replace('=', '').replace('\n', '')
                    # process only known blocks
                    if line.lower() in ExcelBlock:
                        excel_block = line.lower()
                        read_headers = True
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
                    csv_line = next(csv_line)
                    csv_line = [None if col=='' else col for col in csv_line]
                    self.__values[excel_block].append(csv_line)

    @property
    def values(self):
        return self.__values



