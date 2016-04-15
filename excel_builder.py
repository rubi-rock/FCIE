import re
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_range
import ast

from csv_parser import CSVParser
from constants import EXCEL_HEADERS, ExcelBlockDef, ExcelBlock, OUI_NON, OS_LIST, COL_SIZE


def isfloat(value):
  try:
    float(value)
    return True
  except ValueError:
    return False

class ExcelGenerator(object):
    DEFAULT_ROW_HEIGHT = 20
    DEFAULT_ROW_LENGTH = 225

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
        format = ast.literal_eval(str_format)
        return format

    def __reformat_worksheet(self, worksheet, ws_format):
        for name, value in ws_format.items():
            if name == 'orientation':
                if value == 'portrait':
                    worksheet.set_portrait()
                elif value == 'landscape':
                    worksheet.set_landscape()
            elif name == 'paper':
                worksheet.set_paper(int(value))
            elif name == 'tab_color':
                worksheet.tab_color = value
            elif name == 'columns':
                cols = 'A:' + chr(ord('A') + value['count'])
                worksheet.set_column(cols, value['width'])
            elif name == 'tab_color':
                worksheet.column_width = value

    def __process_tab_formatting(self, worksheet, line):
        fmt_worksheet = re.compile('^\s*\%\%({.+})\%\%\n')
        ws_format = fmt_worksheet.split(line.replace(' ', ''))
        if ws_format is not None and len(ws_format) > 1:
            ws_format = ast.literal_eval(ws_format[1])
            self.__reformat_worksheet(worksheet, ws_format)
            return ws_format
        else:
            return None

    def __process_line_formatting(self, line, default_format):
        height = None
        fmt_line = re.compile('^\s*(\{.+\})(.+)$')
        pieces = fmt_line.split(line.replace(' ', ''))
        if pieces is not None and len(pieces) == 4:
            format = ast.literal_eval(pieces[1])
            height = format['heigt'] if 'height' in format else None
            format = self.__workbook.add_format(format)
            line = pieces[2]
        else:
            format = default_format

        return line, format, height

    def __add_textbox(self, worksheet, row, textbox):
        textbox = ast.literal_eval(textbox.replace(' ', ''))
        cell = xl_rowcol_to_cell(row, textbox['start_column'])
        textbox.pop('start_column')
        cell = cell + ':' + xl_rowcol_to_cell(row, textbox['end_column'])
        textbox.pop('end_column')
        #worksheet.merge_range(cell, '', format)
        worksheet.insert_textbox(cell, '', textbox)

    def __add_text_from_file(self, worksheet, row, col, filename, format = None):
        fmt_textbox = re.compile("(.+){\s*'textbox'\s*:(.*)}(.*)")

        if format is None:
            default_format = self.__formats['text']
        else:
            default_format = format

        with open(filename, 'rt') as text_file:
            tab_format = None
            for line in iter(text_file):
                cell_height = None
                tmp = self.__process_tab_formatting(worksheet, line)
                tab_format = tmp if tmp is not None else tab_format
                if tmp is not None:
                    continue

                pieces = fmt_textbox.split(line)
                if pieces is not None and len(pieces) > 1:
                    line = pieces[1]
                    textbox = pieces[2]
                else:
                    textbox = None

                line, format, cell_height = self.__process_line_formatting(line, default_format)

                cell = xl_rowcol_to_cell(row, col)
                cell = cell + ':' + xl_rowcol_to_cell(row, col + tab_format['columns']['count'])
                worksheet.merge_range(cell, line, format)

                if textbox is not None and len(textbox) > 1:
                    self.__add_textbox(worksheet, row, textbox)

                if cell_height is None:
                    cell_height = ExcelGenerator.DEFAULT_ROW_HEIGHT * round(len(line) / tab_format['max_char'])
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
        next_row = self.__add_text_from_file(worksheet, 8, 0, 'frais.txt')
        self.__write_section(worksheet, next_row, 0, EXCEL_HEADERS[ExcelBlock.proposition][ExcelBlockDef.headers])
        next_row += 2
        next_row = self.__add_text_from_file(worksheet, next_row, 0, 'frais_note.txt')
        worksheet.set_h_pagebreaks([next_row])
        next_row += 1
        next_row = self.__add_text_from_file(worksheet, next_row, 0, 'conditions.txt')

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
        worksheet = self.__create_worksheet("Débit préautorisé")
        worksheet.tab_color = "1072BA"
        worksheet.set_column('A:K', self.DEFAULT_ROW_LENGTH / 9)
        self.__add_text_from_file(worksheet, 0, 0, 'debit_form.txt')
