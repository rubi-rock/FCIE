import csv
from io import StringIO
from collections import OrderedDict
from constants import ExcelBlock
from dotmap import DotMap


class CSVParser(object):
    def __init__(self, filename):
        self.__values = DotMap()
        self.__parse(filename)
        self.__compile_data()

    def __compile_data(self):
        self.__create_groups()
        self.__create_user_groups()

    def __create_groups(self):
        self.__values.group = DotMap()
        for institution in self.__values['institution']:
            if institution['numero_groupe'] is None:
                continue
            if institution['numero_groupe'] not in self.__values['group'].keys():
                institution = institution.copy()
                institution.pop('numero_pratique')
                self.__values['group'][institution['numero_groupe']] = institution

    def __get_full_key(self, record):
        #return '{0}.{1}'.format(record.numero_etablissement, record.numero_groupe)
        return '{0}'.format(record.numero_etablissement)

    def __create_user_groups(self):
        # rename
        self.__values.pop('md')
        self.__values.mds = self.__values.pop('proposition')
        self.__values.institutions = self.__values.pop('institution')

        # clean MDs : no group
        for md in self.__values.mds:
            md.pop('numero_groupe')
            md.pop('rmx')

        # Build institution list
        tmp = []
        processed_groups = []
        for institution in self.__values.institutions:
            if institution.numero_groupe is None or institution.numero_etablissement is None:
                continue
            if self.__get_full_key(institution) not in processed_groups:
                institution.full_key = self.__get_full_key(institution)
                cleaned_institution = DotMap(institution.toDict()) # forces a copy because copy() does not work here
                cleaned_institution.pop('numero_pratique')  # remove unrelated data
                tmp.append( cleaned_institution)
                processed_groups.append(cleaned_institution.full_key)
        self.__values.institution_group = self.__values.pop('institutions')
        self.__values.institutions = tmp

        # associate users with institution/groupe
        association_list = DotMap()
        for md in self.__values.mds:
            associations = [None] * len(self.__values.institutions)
            association_list[md.numero] = associations
            idx = 0
            for institution in self.__values.institutions:
                for group in self.__values.institution_group:
                    if md.numero_pratique == group.numero_pratique and institution.full_key == group.full_key and institution.full_key not in associations:
                        associations[idx] = institution.full_key
                idx += 1

            md.institutions = associations
        self.__values.md_institutions = association_list
        self.__values.pop('institution_group')
        self.__values.pop('group')
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
                        self.__values[excel_block].append(DotMap(values))

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
