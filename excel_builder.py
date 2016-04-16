import re
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol
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
        self.__build_document()

#        self.__generate_proposition()
#        self.__generate_config_base()
#        self.__generate_institution()
#        self.__generate_user_access()
#        self.__generate_debit_form()
        self.__workbook.close()

    @staticmethod
    def get_template_name(filename):
        return 'templates/{0}.txt'.format(filename)

    def __write_section(self, worksheet, row, col, section, align='horizontal'):
        for section in section:
            cell = xl_rowcol_to_cell(row, col)
            worksheet.write(cell, section, self.__formats['section'])
            if align == 'horizontal':
                col += 1
            else:
                row += 1

    def __write_empty_cells(self, worksheet, row, col, section, align='horizontal'):
        for offset in range[len(section)]:
            if align == 'horizontal':
                cell = xl_rowcol_to_cell(row, col + offset)
            else:
                cell = xl_rowcol_to_cell(row + offset, col)
            worksheet.write(cell, None, self.__formats['cell'])

    def __write_headers(self, worksheet, row, col, headers, format=None):
        if format is None:
            format = self.__formats['header']
        for header in headers:
            cell = xl_rowcol_to_cell(row, col)
            worksheet.write(cell, header, format)
            col += 1

    def __write_cell(self, worksheet, row, col, value, format=None):
        if format is None:
            format = self.__formats['cell']
        cell = xl_rowcol_to_cell(row, col)
        worksheet.write(cell, value, format)

        sizes = COL_SIZE[block]
        col_num = 0
        for size in sizes:
            worksheet.set_column(col_num, col_num, size)
            col_num += 1

    @staticmethod
    def cell_as_int(cell_value):
        return int(cell_value) if cell_value is not None else None

    @staticmethod
    def cell_as_oui_non(cell_value):
        value = OUI_NON[cell_value]
        return value

    @staticmethod
    def call_as_percentage(cell_value):
        return float(cell_value.strip('%')) / 100 if cell_value is not None else None

    def __build_document(self):
        template = self.get_template_name('document_template')
        with open(template, 'rt') as text_file:
            for line in iter(text_file):
                if line.strip() == '':
                    continue
                params = ast.literal_eval(line)
                if 'style' in params.keys():
                    self.__add_format(params['style'])
                elif 'tab' in params.keys():
                    self.__build_worksheet(params['tab'])

    def __add_format(self, style):
        name = style['name']
        properties = style['properties']
        self.__formats[name] = self.__workbook.add_format(properties)

    def __create_worksheet(self, name):
        worksheet = self.__workbook.add_worksheet(name)
        worksheet.set_landscape()
        worksheet.set_paper(5)
        worksheet.set_margins(0.5, 0.5, 1.5, 0.8)
        worksheet.set_header('&L&G&R&F', {'image_left': 'purkinje.png'})
        worksheet.set_footer('&L&A&CPage &P / &N&R&D')
        return worksheet

    @staticmethod
    def __reformat_worksheet(worksheet, ws_format):
        for name, value in ws_format.items():
            if name == 'orientation':
                if value == 'portrait':
                    worksheet.set_portrait()
                elif value == 'landscape':
                    worksheet.set_landscape()
            elif name == 'paper':
                worksheet.set_paper(int(value))
            elif name == 'margins':
                worksheet.set_margins(value)
            elif name == 'tab_color':
                worksheet.tab_color = value
            elif name == 'columns':
                if type(value['width']) is int:
                    cols = 'A:' + chr(ord('A') + value['count'])
                    worksheet.set_column(cols, value['width'])
                elif type(value['width']) is list:
                    i = 0
                    for col_width in value['width']:
                        col_cel = chr(ord('A') + i)
                        col_cel = col_cel + ':' + col_cel
                        worksheet.set_column(col_cel, col_width)
                        i += 1

            elif name == 'tab_color':
                worksheet.column_width = value

    def __instanciate_worksheet(self, ws_definition):
        worksheet = self.__workbook.add_worksheet(ws_definition['name'])
        self.__reformat_worksheet(worksheet, ws_definition['format'])
        worksheet.set_margins(0.5, 0.5, 1.5, 0.8)
        if ws_definition['header'] is not None and ws_definition['header']['format'] is not None:
            worksheet.set_header(ws_definition['header']['format'], ws_definition['header']['options'])
        if ws_definition['footer'] is not None and ws_definition['footer']['format'] is not None:
            worksheet.set_footer(ws_definition['footer']['format'], ws_definition['footer']['options'] if 'option' in ws_definition['footer'].keys() else None)
        return worksheet

    def __build_worksheet(self, filename):
        ws_definition = { 'name': '',  'format': {'format': None, 'options': None}, 'content': [], 'header': {'format': None, 'options': None}, 'footer': {}}
        with open(self.get_template_name(filename), 'rt') as text_file:
            for line in iter(text_file):
                if line.strip() == '':
                    continue
                params = ast.literal_eval(line)
                if 'name' in params.keys():
                    ws_definition['name'] = params['name']
                    params.pop('name')
                    ws_definition['format'] = params
                elif 'section' in params.keys():
                    ws_definition['content'].append(params)
                elif 'header' in params.keys():
                    ws_definition['header'] = params['header']
                elif 'footer' in params.keys():
                    ws_definition['footer'] = params['footer']
                elif 'cell' in params.keys():
                    ws_definition['content'].append(params)
                elif 'col' in params.keys():
                    ws_definition['content'].append(params)
                elif 'row' in params.keys():
                    ws_definition['content'].append(params)

        worksheet =  self.__instanciate_worksheet(ws_definition)

        last_row, last_column = 0, 0
        for item in ws_definition['content']:
            if 'cell' in item.keys():
                lr, lc = self.__fill_cell(worksheet, item, last_row, last_column)
            elif 'col' in item.keys():
                lr, lc = self.__fill_col(worksheet, item, last_row, last_column)
            elif 'row' in item.keys():
                lr, lc = self.__fill_row(worksheet, item, last_row, last_column)
            if 'remember_last_row' in item:
                last_row = lr
            if 'remember_last_column' in item:
                last_column = lc

        return worksheet

    @staticmethod
    def __substitute_last_coords(cell, last_row, last_column):
        pieces = cell.split(':')
        if len(pieces) == 1:
            return cell

        if 'last_column' == pieces[0]:
            py_code = pieces[0]
            pieces[0] = chr(ord('A') + eval(py_code))

        if 'last_row' in pieces[1]:
            py_code = pieces[1]
            pieces[1] = str(eval(py_code))

        cell = pieces[0]  + pieces[1]
        return cell

    @staticmethod
    def __get_casted_value(value, definition):
        if value is None:
            return value

        if 'cast' in definition:
            obj_method = getattr(ExcelGenerator, definition['cast'])
            if obj_method is not None and callable(obj_method):
                return obj_method(value)
        return value

    def __fill_cell(self, worksheet, item, last_row, last_column):
        cell = item['cell']
        cell = self.__substitute_last_coords(cell, last_row, last_column)
        style = item['style'] if 'style' in item else None
        if type(style) is str:
            format = self.__formats[style]
        else:
            format = self.__workbook.add_format(style)
        value = item['value'] if 'value' in item else None
        worksheet.write(cell, self.__get_casted_value(value, item), format)
        row, col = xl_cell_to_rowcol(cell)
        return row, col

    def __fill_col(self, worksheet, item, last_row, last_column):
        col = item['col']
        col = self.__substitute_last_coords(col, last_row, last_column)
        row, col = xl_cell_to_rowcol(col)
        style = item['style'] if 'style' in item else None
        if type(style) is str:
            format = self.__formats[style]
        else:
            format = self.__workbook.add_format(style)

        if 'loop' in item.keys():
            value_list = self.__csv.values[item['loop']]
        else:
            value_list = None

        row_offset = 0
        if value_list is not None:
            idx = item['index'] if 'index' in item.keys() else None
            if idx is None:
                value = item['value'] if 'value' in item.keys() else None
            for csv_values in value_list:
                cell = xl_rowcol_to_cell(row + row_offset, col)
                if idx is not None:
                    value = csv_values[idx]
                worksheet.write(cell, self.__get_casted_value(value, item), format)
                if 'validation' in item.keys():
                    worksheet.data_validation(cell, item['validation'].copy())
                row_offset += 1
        else:
            for value in item['values']:
                cell = xl_rowcol_to_cell(row, col + row_offset)
                worksheet.write(cell, self.__get_casted_value(value, item), format)
                if 'validation' in item.keys():
                    worksheet.data_validation(cell, item['validation'].copy())
                row_offset += 1

        if 'spare_rows' in item:
            for i in range(item['spare_rows']):
                cell = xl_rowcol_to_cell(row + row_offset, col)
                worksheet.write(cell, None, format)
                if 'validation' in item.keys():
                    worksheet.data_validation(cell, item['validation'].copy())
                row_offset += 1

        return row + row_offset, col

    def __fill_row(self, worksheet, item, last_row, last_column):
        row = item['row']
        row = self.__substitute_last_coords(row, last_row, last_column)
        row, col = xl_cell_to_rowcol(row)
        style = item['style'] if 'style' in item else None
        if type(style) is str:
            format = self.__formats[style]
        else:
            format = self.__workbook.add_format(style)

        if 'loop' in item.keys():
            value_list = self.__csv.values[item['loop']]
        else:
            value_list = None

        row_offset = 0
        col_ofsset = 0
        if value_list is not None:
            for csv_values in value_list:
                for idx in item['indexes']:
                    cell = xl_rowcol_to_cell(row + row_offset, col + col_ofsset)
                    value = csv_values[idx] if idx is not None else None
                    worksheet.write(cell, self.__get_casted_value(value, item), format)
                    if 'validation' in item.keys():
                        worksheet.data_validation(cell, item['validation'].copy())
                    col_ofsset += 1
                row_offset += 1
                col_ofsset = 0
        else:
            for value in item['values']:
                cell = xl_rowcol_to_cell(row, col + col_ofsset)
                worksheet.write(cell, self.__get_casted_value(value, item), format)
                if 'validation' in item.keys():
                    worksheet.data_validation(cell, item['validation'].copy())
                col_ofsset += 1

        return row + row_offset, col + col_ofsset

    def __add_text_lines(self, worksheet, row, col, col_count, lines, format = None):
        if format is None:
            format = self.__formats['text']
        for line in lines:
            cell = xl_rowcol_to_cell(row, col)
            cell = cell + ':' + xl_rowcol_to_cell(row, col + col_count)
            worksheet.merge_range(cell, line, format)
            row += 1
        return row

    @staticmethod
    def __parse_format(str_format):
        format = ast.literal_eval(str_format)
        return format

    def __process_tab_formatting(self, worksheet, line):
        fmt_worksheet = re.compile('^\s*%%({.+})%%\n')
        ws_format = fmt_worksheet.split(line.replace(' ', ''))
        if ws_format is not None and len(ws_format) > 1:
            ws_format = ast.literal_eval(ws_format[1])
            self.__reformat_worksheet(worksheet, ws_format)
            return ws_format
        else:
            return None

    def __process_line_formatting(self, line, default_format):
        height = None
        pieces = ast.literal_eval(line)
        fmt_line = re.compile('^\s*(\{.+\})\s*(.+)$')
        pieces = fmt_line.split(line)
        if pieces is not None and len(pieces) == 4:
            format = ast.literal_eval(pieces[1])
            if 'height' in format:
                height = format['height']
                format.pop('height')
            else:
                height = None
            format = self.__workbook.add_format(format)
            line = pieces[2]
        else:
            format = default_format

        return line, format, height

    @staticmethod
    def __add_textbox(worksheet, row, textbox):
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

        with open(self.get_template_name(filename), 'rt') as text_file:
            tab_format = None
            for line in iter(text_file):
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
        worksheet = self.__build_worksheet('proposition_tab')
        # Customer
        self.__write_section(worksheet, 0, 0, EXCEL_HEADERS[ExcelBlock.attention_de][ExcelBlockDef.headers], 'vertical')
        self.__write_empty_cells(worksheet, 0, 1, EXCEL_HEADERS[ExcelBlock.attention_de][ExcelBlockDef.headers], 'vertical')
        # Purkinje Agent
        self.__write_section(worksheet, 0, 5, EXCEL_HEADERS[ExcelBlock.agent_purkinje][ExcelBlockDef.headers], 'vertical')
        self.__write_empty_cells(worksheet, 0, 6, EXCEL_HEADERS[ExcelBlock.agent_purkinje][ExcelBlockDef.headers], 'vertical')
        # Fees
        next_row = self.__add_text_from_file(worksheet, 8, 0, 'frais')
        self.__write_section(worksheet, next_row, 0, EXCEL_HEADERS[ExcelBlock.proposition][ExcelBlockDef.headers])
        next_row += 2
        next_row = self.__add_text_from_file(worksheet, next_row, 0, 'frais_note')
        worksheet.set_h_pagebreaks([next_row])
        next_row += 1
        self.__add_text_from_file(worksheet, next_row, 0, 'conditions')

    def __generate_config_base(self):
        worksheet = self.__build_worksheet('config_base')
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
            self.__write_cell(worksheet, row, 0, self.cell_as_int(md[0]))  # Numero de pratique
            self.__write_cell(worksheet, row, 1, self.cell_as_int(md[1]))  # Numero de groupe
            self.__write_cell(worksheet, row, 2, md[2])  # Nom
            self.__write_cell(worksheet, row, 3, md[3])  # Prénon
            self.__write_cell(worksheet, row, 4, md[4])  # Specialité
            self.__write_cell(worksheet, row, 5, OUI_NON[md[5]])  # RMX (oui/non)
            self.__write_cell(worksheet, row, 6, None)  # Inc (oui/non)
            self.__write_cell(worksheet, row, 7, None)  # Nom compagnie
            self.__write_cell(worksheet, row, 8, None)  # Date fin année fiscalle inc
            row += 1

    def __generate_institution(self):
        worksheet = self.__build_worksheet('institution_tab')
        self.__write_section(worksheet, 0, 0, EXCEL_HEADERS[ExcelBlock.institution][ExcelBlockDef.section])
        self.__write_headers(worksheet, 1, 0, EXCEL_HEADERS[ExcelBlock.institution][ExcelBlockDef.headers])
        worksheet.set_column('A:H', 15)
        worksheet.set_column('I:I', 25)
        # Etablissement
        mds = self.__csv.values[ExcelBlock.institution]
        row = 2
        for md in mds:
            self.__write_cell(worksheet, row, 0, self.cell_as_int(md[0]))  # Numero de pratique
            self.__write_cell(worksheet, row, 1, self.cell_as_int(md[1]))  # Numero de groupe
            self.__write_cell(worksheet, row, 2, self.cell_as_int(md[2]))  # Numero d'etblissement
            self.__write_cell(worksheet, row, 3, md[3])  # Nom d'etablissement
            self.__write_cell(worksheet, row, 4, float(md[4].strip('%'))/100 if md[4] is not None else None, self.__formats['cell_percent'])  # Pourcentage
            self.__write_cell(worksheet, row, 5, OUI_NON[md[5]])  # RMX (oui/non)
            self.__write_cell(worksheet, row, 6, OUI_NON[md[6]])  # Secteur Cabinet (oui/non)
            self.__write_cell(worksheet, row, 7, None)  # Secteur CLSC
            self.__write_cell(worksheet, row, 8, None)  # Secteur Centre Hospitalier
            row += 1

    def __generate_user_access(self):
        worksheet = self.__build_worksheet('user_access_tab')
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
        self.__build_worksheet('debit_tab')
