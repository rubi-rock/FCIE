import re
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
import pyexcel
import pyexcel_xlsx

from csv_parser import CSVParser
from constants import EXCEL_HEADERS, ExcelBlockDef, ExcelBlock, OUI_NON, OS_LIST, COL_SIZE

class ExcelGenerator(object):
    DEFAULT_ROW_HEIGHT = 20
    DEFAULT_ROW_LENGHT = 225

    def __init__(self, filename):
        self.__csv = CSVParser(filename)
        self.__filename = filename.replace('csv', 'xlsx')
        self.__workbook = xlsxwriter.Workbook(self.__filename)
        self.__OS_list = '\n'.join(['[ ] ' + os for os in OS_LIST])

        self.__formats = {}
        self.__init_formats()

        self.__generate_proposition()
        self.__generate_config_base()
        self.__generate_institution()
        self.__generate_user_access()
        self.__generate_debit_form()
        self.__workbook.close()

    def __write_section(self, worksheet, row, col, section, align = 'horizontal'):
        for section in section:
            cell = xl_rowcol_to_cell(row, col)
            worksheet.write(cell, section, self.__formats['section'])
            if align == 'horizontal':
                col += 1
            else:
                row += 1

    def __write_empty_cells(self, worksheet, row, col, section, align = 'horizontal'):
        for section in section:
            cell = xl_rowcol_to_cell(row, col)
            worksheet.write(cell, None, self.__formats['cell'])
            if align == 'horizontal':
                col += 1
            else:
                row += 1

    def __write_headers(self, worksheet, row, col, headers, format = None):
        if format is None:
            format = self.__formats['header']
        for header in headers:
            cell = xl_rowcol_to_cell(row, col)
            worksheet.write(cell, header, format)
            col += 1

    def __write_cell(self, worksheet, row, col, value, format = None):
        if format is None:
            format = self.__formats['cell']
        cell = xl_rowcol_to_cell(row, col)
        worksheet.write(cell, value, format)

    def __set_col_sizes(self, worksheet, block):
        if block not in COL_SIZE.keys():
            return

        sizes = COL_SIZE[block]
        col_num = 0
        for size in sizes:
            worksheet.set_column(col_num, col_num, size)
            col_num += 1

    def __cell_as_int(self, cell_value):
        return int(cell_value) if cell_value is not None else None

    def __init_formats(self):
        self.__formats['section'] = self.__workbook.add_format(
            {'border': 6, 'bold': True, 'font_name': 'Times New Roman'})
        self.__formats['header'] = self.__workbook.add_format(
            {'border': 6, 'bold': False, 'font_name': 'Times New Roman'})
        self.__formats['cell'] = self.__workbook.add_format({'valign': 'top', 'text_wrap': True, 'border': 1})
        self.__formats['cell_percent'] = self.__workbook.add_format({'valign': 'top', 'num_format': '0.0%', 'text_wrap': True, 'border': 1})
        self.__formats['text'] = self.__workbook.add_format({'valign': 'top', 'text_wrap': True, 'font_name': 'Book Antiqua', 'font_size': '11'})

    def __create_worksheet(self, name):
        worksheet = self.__workbook.add_worksheet(name)
        worksheet.set_landscape()
        worksheet.set_paper(5)
        worksheet.set_margins(0.5, 0.5, 1.5, 0.8)
        worksheet.set_header('&L&G&R&F', {'image_left': 'purkinje.png'})
        worksheet.set_footer('&L&A&CPage &P / &N&R&D')
        return worksheet

    def __add_text_lines(self, worksheet, row, col, col_count, lines, format = None):
        if format is None:
            format = self.__formats['text']
        for line in lines:
            cell = xl_rowcol_to_cell(row, col)
            cell = cell + ':' + xl_rowcol_to_cell(row, col + col_count)
            worksheet.merge_range(cell, line, format)
            row += 1
        return row

    def __parse_format(self, str_format):
        format = {}
        pieces = str_format.replace('{', '').replace('}', '').replace(' ', '').split(',')
        for piece in pieces:
            name, value = piece.split(':')
            if value == 'True':
                value = True
            elif value == 'False':
                value = False
            elif value.isdigit():
                value = int(value)
            elif value.isfloat():
                value = float(value)
            format[name] = value
        return format

    def __add_text_from_file(self, worksheet, row, col, col_count, filename, format = None):
        fmt_regex = re.compile('(\{.+\})(.+)\n')
        if format is None:
            default_format = self.__formats['text']
        else:
            default_format = format
        with open(filename, 'rt') as text_file:
            for line in iter(text_file):
                pieces = fmt_regex.split(line)
                if len(pieces) == 4:
                    format = self.__parse_format(pieces[1])
                    format = self.__workbook.add_format(format)
                    line = pieces[2]
                else:
                    format = default_format

                cell = xl_rowcol_to_cell(row, col)
                cell = cell + ':' + xl_rowcol_to_cell(row, col + col_count)
                worksheet.merge_range(cell, line, format)

                cell_height = ExcelGenerator.DEFAULT_ROW_HEIGHT * round(len(line)/ExcelGenerator.DEFAULT_ROW_LENGHT)
                if cell_height < ExcelGenerator.DEFAULT_ROW_HEIGHT:
                    cell_height = ExcelGenerator.DEFAULT_ROW_HEIGHT
                worksheet.set_row(row, cell_height)
                row += 1
        return row

    def __generate_proposition(self):
        worksheet = self.__create_worksheet("Proposition Financière")
        worksheet.tab_color = "663300"
        self.__set_col_sizes(worksheet, ExcelBlock.proposition)
        # Customer
        self.__write_section(worksheet, 0, 0, EXCEL_HEADERS[ExcelBlock.attention_de][ExcelBlockDef.headers], 'vertical')
        self.__write_empty_cells(worksheet, 0, 1, EXCEL_HEADERS[ExcelBlock.attention_de][ExcelBlockDef.headers], 'vertical')
        # Purkinje Agent
        self.__write_section(worksheet, 0, 5, EXCEL_HEADERS[ExcelBlock.agent_purkinje][ExcelBlockDef.headers], 'vertical')
        self.__write_empty_cells(worksheet, 0, 6, EXCEL_HEADERS[ExcelBlock.agent_purkinje][ExcelBlockDef.headers], 'vertical')
        # Fees
        next_row = self.__add_text_from_file(worksheet, 8, 0, len(EXCEL_HEADERS[ExcelBlock.proposition][ExcelBlockDef.headers]), 'frais.txt')
        self.__write_section(worksheet, next_row, 0, EXCEL_HEADERS[ExcelBlock.proposition][ExcelBlockDef.headers])
        next_row += 2
        next_row = self.__add_text_from_file(worksheet, next_row, 0, len(EXCEL_HEADERS[ExcelBlock.proposition][ExcelBlockDef.headers]), 'frais_note.txt')
        worksheet.set_h_pagebreaks([next_row])
        next_row += 1
        next_row = self.__add_text_from_file(worksheet, next_row, 0, len(EXCEL_HEADERS[ExcelBlock.proposition][ExcelBlockDef.headers]), 'conditions.txt')


    def __generate_config_base(self):
        worksheet = self.__create_worksheet("Config de base")
        worksheet.tab_color = "FF3300"
        self.__set_col_sizes(worksheet, ExcelBlock.customer)
        self.__write_section(worksheet, 0, 0, EXCEL_HEADERS[ExcelBlock.customer][ExcelBlockDef.section])
        self.__write_headers(worksheet, 1, 0, EXCEL_HEADERS[ExcelBlock.customer][ExcelBlockDef.headers])
        self.__write_section(worksheet, 4, 0, EXCEL_HEADERS[ExcelBlock.user][ExcelBlockDef.section])
        self.__write_headers(worksheet, 5, 0, EXCEL_HEADERS[ExcelBlock.user][ExcelBlockDef.headers])
        #Customers
        customer = self.__csv.values[ExcelBlock.customer][0]
        self.__write_cell(worksheet, 2, 0, None)    # Nom du client / agence
        self.__write_cell(worksheet, 2, 1, customer[0])   # Numéro agence
        self.__write_cell(worksheet, 2, 2, customer[1])   # Mot de passe TIP-I
        self.__write_cell(worksheet, 2, 3, None)  # Ville
        # Users
        users = self.__csv.values[ExcelBlock.user]
        row = 6
        for user in users:
            self.__write_cell(worksheet, row, 0, user[0])  # Nom utilisateur (username)
            self.__write_cell(worksheet, row, 1, user[1])  # Mot de passe
            self.__write_cell(worksheet, row, 2, user[2])  # Nom
            self.__write_cell(worksheet, row, 3, None)  # Prenom
            self.__write_cell(worksheet, row, 4, self.__OS_list)
            row += 1
        row += 1
        self.__write_section(worksheet, row, 0, EXCEL_HEADERS[ExcelBlock.md][ExcelBlockDef.section])
        row += 1
        self.__write_headers(worksheet, row, 0, EXCEL_HEADERS[ExcelBlock.md][ExcelBlockDef.headers])
        row += 1
        # MD
        mds = self.__csv.values[ExcelBlock.md]
        #row = 6
        for md in mds:
            self.__write_cell(worksheet, row, 0, self.__cell_as_int(md[0]))  # Numero de pratique
            self.__write_cell(worksheet, row, 1, self.__cell_as_int(md[1]))  # Numero de groupe
            self.__write_cell(worksheet, row, 2, md[2])  # Nom
            self.__write_cell(worksheet, row, 3, md[3])  # Prénon
            self.__write_cell(worksheet, row, 4, md[4])  # Specialité
            self.__write_cell(worksheet, row, 5, OUI_NON[md[5]])  # RMX (oui/non)
            self.__write_cell(worksheet, row, 6, None)  # Inc (oui/non)
            self.__write_cell(worksheet, row, 7, None)  # Nom compagnie
            self.__write_cell(worksheet, row, 8, None)  # Date fin année fiscalle inc
            row += 1

    def __generate_institution(self):
        worksheet = self.__create_worksheet("Établissement")
        worksheet.tab_color = "96FF33"
        self.__set_col_sizes(worksheet, ExcelBlock.institution)
        self.__write_section(worksheet, 0, 0, EXCEL_HEADERS[ExcelBlock.institution][ExcelBlockDef.section])
        self.__write_headers(worksheet, 1, 0, EXCEL_HEADERS[ExcelBlock.institution][ExcelBlockDef.headers])
        worksheet.set_column('A:H', 15)
        worksheet.set_column('I:I', 25)
        # Etablissement
        mds = self.__csv.values[ExcelBlock.institution]
        row = 2
        for md in mds:
            self.__write_cell(worksheet, row, 0, self.__cell_as_int(md[0]))  # Numero de pratique
            self.__write_cell(worksheet, row, 1, self.__cell_as_int(md[1]))  # Numero de groupe
            self.__write_cell(worksheet, row, 2, self.__cell_as_int(md[2]))  # Numero d'etblissement
            self.__write_cell(worksheet, row, 3, md[3])  # Nom d'etablissement
            self.__write_cell(worksheet, row, 4, float(md[4].strip('%'))/100 if md[4] != None else None, self.__formats['cell_percent'])  # Pourcentage
            self.__write_cell(worksheet, row, 5, OUI_NON[md[5]])  # RMX (oui/non)
            self.__write_cell(worksheet, row, 6, OUI_NON[md[6]])  # Secteur Cabinet (oui/non)
            self.__write_cell(worksheet, row, 7, None)  # Secteur CLSC
            self.__write_cell(worksheet, row, 8, None)  # Secteur Centre Hospitalier
            row += 1

    def __generate_user_access(self):
        worksheet = self.__create_worksheet("Accès utilisateurs")
        worksheet.tab_color = "1072BA"
        self.__set_col_sizes(worksheet, ExcelBlock.user_access)
        self.__write_section(worksheet, 0, 0, EXCEL_HEADERS[ExcelBlock.user_access][ExcelBlockDef.section])
        self.__write_headers(worksheet, 1, 0, EXCEL_HEADERS[ExcelBlock.user_access][ExcelBlockDef.headers])
        # User Access
        user_list = self.__csv.values[ExcelBlock.user_access]
        row = 2
        for user in user_list:
            self.__write_cell(worksheet, row, 0, user[0])  # Utilisateur
            self.__write_cell(worksheet, row, 1, user[1])  # Group
            row += 1

    def __generate_debit_form(self):
        pass
