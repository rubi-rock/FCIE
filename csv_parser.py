import csv
from io import StringIO
from constants import ExcelBlock


class CSVParser(object):
    def __init__(self, filename):
        self.__values = {}
        self.__parse(filename)
        self.__build_tree()

    def __build_tree(self):
        pass

    def __parse(self, filename):
        excel_block = None
        read_headers = False
        with open(filename, mode='rt', newline="\n") as csv_file:
            for line in iter(csv_file):
                # block detection
                if line.startswith("==="):
                    line = line.replace('=', '').replace('\n', '')
                    # process only known blocks
                    if line.lower() in ExcelBlock:
                        excel_block = line.lower()
                        read_headers = True
                        if excel_block not in ['site', 'customer']:
                            self.__values[excel_block] = []
                        continue
                    else:
                        excel_block = None


                # skip headers
                if read_headers:
                    csv_stream = StringIO(line)
                    csv_line = csv.reader(csv_stream, delimiter=',', quotechar='"')
                    csv_line = [None if col == '' else col for col in csv_line]
                    headers = list(None if header == '' else header.lower().replace(' ', '_') for header in csv_line[0])
                    read_headers = False
                    continue

                # add values to block
                if excel_block is not None:
                    csv_stream = StringIO(line)
                    csv_line = csv.reader(csv_stream, delimiter=',', quotechar='"')
                    csv_line = next(csv_line)
                    csv_line = [None if col == '' else col for col in csv_line]
                    values = dict(zip(headers, csv_line))
                    if excel_block in ['site', 'customer']:
                        self.__values[excel_block] = values
                    else:
                        self.__values[excel_block].append(values)

    @property
    def values(self):
        return self.__values

    def has_key(self, key):
        keys = key.split('.')
        d = self.__values
        for k in keys:
            if k in d:
                d = d[k]
            else:
                return False
        return True

    def get_value(self, key):
        keys = key.split('.')
        d = self.__values
        for k in keys:
            if k in d:
                d = d[k]
            else:
                return None
        return d

    def get_column(self, key):
        keys = key.split('.')
        if len(keys) < 2:
            return None

        record_name = keys[0]
        if keys[0] in self.__values.keys():
            record_list = self.__values[record_name]
            field_name = keys[1]
            if len(record_list) > 0:
                if field_name in record_list[0].keys():
                    return [record[field_name] for record in record_list]
        return None

    def get_record(self, key):
        keys = key.split('.')
        if len(keys) > 1:
            keys = keys[0]
        return self.get_value('.'.join(keys))

    def get_list_length(self, key):
        keys = key.split('.')
        if len(keys) > 0:
            return len(self.__values[keys[0]])

        return 0
