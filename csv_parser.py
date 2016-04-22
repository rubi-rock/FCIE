import csv
from io import StringIO
from collections import Counter
from constants import ExcelBlock
from dotmap import DotMap
import logging


class CSVParser(object):
    def __init__(self, filename):
        self.__values = DotMap()
        self.__parse(filename)
        self.__compile_data()

    def __compile_data(self):
        self.__copy_mdusers_to_proposition()
        self.__cleanup_names()
        self.__adjust_MDs()
        self.__adjust_Users()
        self.__create_user_groups()
        self.__update_MD_biller_status()
        self.__create_user_md_association()
        self.__create_md_user_association()
        self.__log_data()

    def __log_DotMap(self, name, obj):
        logging.info(name)
        if type(obj) is DotMap:
            for value in obj.items():
                logging.info('\t' + str(value))
        elif type(obj) is list:
            for value in obj:
                logging.info('\t' + str(value))

    def __log_data(self):
        self.__log_DotMap('site', self.__values.site)
        self.__log_DotMap('mds', self.__values.mds)
        self.__log_DotMap('users', self.__values.users)
        self.__log_DotMap('md_user', self.__values.md_user)
        self.__log_DotMap('user_md', self.__values.user_md)
        self.__log_DotMap('user_institutions', self.__values.user_institutions)

    def __create_institutions(self):
        self.__values.institution_group = self.__values.pop('institution')
        self.__values.institutions = DotMap()
        for group in self.__values.institution_group:
            if group.numero_groupe is None:
                continue
            group.full_key = self.__get_full_key(group)
            if group.numero_etablissement not in self.__values.institutions.keys():
                institution = DotMap(group.toDict()) # forces a copy because copy() does not work here
                institution.pop('numero_pratique')
                institution.id = '{0} ({1})'.format(institution.nom_etablissement, institution.numero_etablissement)
                self.__values.institutions[institution.numero_etablissement] = institution
        self.__values.institutions = list(self.__values.institutions.values())

    def __get_full_key(self, record):
        #return '{0}.{1}'.format(record.numero_etablissement, record.numero_groupe)
        return '{0}'.format(record.numero_etablissement)

    def __cleanup_names(self):
        # rename
        #self.__values.pop('md')
        self.__values.mds = self.__values.pop('md')
        self.__create_institutions()
        self.__values.users = self.__values.pop('user')

    def __copy_mdusers_to_proposition(self):
        for md in self.__values.md:
            if 'users' not in md.keys():
                continue
            for prop in self.__values.proposition:
                if md.numero_pratique == prop.numero:
                    prop['users'] = md.users
        pass

    def __adjust_MDs(self):
        # clean MDs : no group
        for md in self.__values.mds:
            md.prenom = md.prenom.upper()
            spec = md.specialite.replace('/', '\n').split(' ', 1)
            md.code_specilite = spec[0]
            md.specialite = spec[1]
            md.nom = md.nom.upper()
            md.is_biller = True
            md.id = '{0}, {1} ({2})'.format(md.prenom, md.nom, md.numero if 'numero' in md.keys() else '')
            if 'users' in md.keys():
                md.users = md.users.split('|')
        pass

    def __adjust_Users(self):
        for user in self.__values.users:
            user.id = '{0}, {1} ({2})'.format(user.prenom, user.nom, user.utilisateur if 'utilisateur' in user.keys() else '')
        pass

    def __create_user_groups(self):
        # associate users with institution/groupe
        association_list = DotMap()
        for md in self.__values.mds:
            associations = [None] * len(self.__values.institutions)
            association_list[md.numero_pratique] = associations
            idx = 0
            for institution in self.__values.institutions:
                for group in self.__values.institution_group:
                    group.full_key = self.__get_full_key(group)
                    if md.numero_pratique == group.numero_pratique and institution.full_key == group.full_key and institution.full_key not in associations:
                        associations[idx] = institution.full_key
                idx += 1

            md.institutions = associations.copy()
        self.__values.md_institutions = list(association_list.values())
        self.__values.pop('institution_group')
        self.__values.pop('group')

    def __update_MD_biller_status(self):
        billers = self.__values.mds
        return

        billers = {}
        for md in self.__values.mds:
            for institution in md.institutions:
                biller = institution is not None
                if biller:
                    md.is_biller = True
                    if md.numero not in billers.keys():
                        billers[md.numero] = md
            continue
        self.__values.billers = list(billers.values())

    def __create_user_md_association(self):
        association_list = DotMap()
        for user in self.__values.users:
            associations = [None] * len(self.__values.mds)
            idx = 0
            for md in self.__values.mds:
                if 'users' not in md.keys():
                    continue
                for link in md.users:
                    if user.utilisateur == link:
                        associations[idx] =  md.numero_groupe
                idx += 1
            # Check users without limitations, then it means they can bill for every md
            counts = Counter(associations)
            if len(counts) == 1 and counts[None] == len(self.__values.mds):
                associations = []
                for md in self.__values.mds:
                    associations.append(md.numero_groupe)

            association_list[user.utilisateur] = associations

        self.__values.user_md = list(association_list.values())
        pass

    def __create_md_user_association(self):
        association_list = DotMap()
        for md in self.__values.mds:
            associations = [None] * len(self.__values.users)
            idx = 0
            for user in self.__values.users:
                if 'users' not in md.keys():
                    continue
                for link in md.users:
                    if user.utilisateur == link:
                        associations[idx] = user.utilisateur
                idx += 1
            # Check users without limitations, then it means they can bill for every md
            counts = Counter(associations)
            if len(counts) == 1 and counts[None] == len(self.__values.users):
                associations = []
                for user in self.__values.users:
                    associations.append(user.utilisateur if md.is_biller else None)

            association_list[md.numero_pratique] = associations

        self.__values.md_user = list(association_list.values())
        pass

    def __parse(self, filename):
        excel_block = None
        read_headers = False
        with open(filename, mode='rt', encoding='iso-8859-1', newline="\r\n") as csv_file:
            for line in iter(csv_file):
                line = line.replace('\r', '').replace('\n', '')
                # block detection
                if line.startswith("==="):
                    line = line.replace('=', '').replace('\n', '')
                    # process only known blocks
                    if line.lower().strip() in ExcelBlock:
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
                    csv_line = [True if col == 'T' else col for col in csv_line]
                    csv_line = [False if col == 'F' else col for col in csv_line]
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
        return self.get_value(keys[0])

    def get_list_length(self, key):
        keys = key.split('.')
        if len(keys) > 0:
            return len(self.__values[keys[0]])

        return 0

