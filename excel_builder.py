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
        self.__variables = {}
        self.__formats = {}
        self.__build_document()
        self.__workbook.close()

    @property
    def variables(self):
        return self.__variables

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
            elif name == 'hide_gridlines':
                worksheet.hide_gridlines(value)
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
                line = line.strip()
                if line == '' or line.startswith('#'):
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

        self.__variables.clear()
        self.__variables['last_row'], self.__variables['last_column'] = 0, 0
        for item in ws_definition['content']:
            if 'cell' in item.keys():
                lr, lc = self.__fill_cell(worksheet, item)
            elif 'col' in item.keys():
                lr, lc = self.__fill_col(worksheet, item)
            elif 'row' in item.keys():
                lr, lc = self.__fill_row(worksheet, item)

            if 'remember_last_row' in item:
                self.__variables['last_row'] = lr + 1
            if 'remember_last_column' in item:
                self.__variables['last_column'] = lc + 1

        return worksheet

    def __prepare_eval_expression(self, py_code):
        if py_code is None or type(py_code) is not str :
            return
        for var_name in self.__variables.keys():
            if var_name in py_code:
                py_code = py_code.replace( var_name, 'self.variables["' + var_name + '"]')
        return py_code

    def __substitute_variables(self, cell_value):
        if cell_value is None or type(cell_value) is not str :
            return
        for var_name in self.__variables.keys():
            if var_name in cell_value:
                cell_value = cell_value.replace(var_name, str(self.variables[var_name]))
        return cell_value.replace('~', '')

    def __substitute_last_coords(self, cell):
        pieces = cell.split('~')
        if len(pieces) == 1:
            return cell

        coord_components = []
        for piece in pieces:
            if ':' not in piece:
                coord_components.append(piece)
            else:
                for tmp in re.split('(^[^:]+)(:)([^:]+$)', piece):
                    if len(tmp) > 0:
                        coord_components.append(tmp)

        for idx in range(len(coord_components)):
            if 'last_column' in coord_components[idx]:
                py_code = self.__prepare_eval_expression(coord_components[idx])
                coord_components[idx] = chr(ord('A') + eval(py_code))

            if 'last_row' in coord_components[idx]:
                py_code = self.__prepare_eval_expression(coord_components[idx])
                coord_components[idx] = str(eval(py_code))

        cell = ''.join(coord_components)
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

    def write_value(self, worksheet, item, cell_value, format, row, col, row_offset, merge_to_row, merge_to_col):
        cell = xl_rowcol_to_cell(row + row_offset, col)
        self.__variables['current_row'] = row + row_offset + 1
        cell_value = self.__substitute_variables(cell_value)

        if merge_to_col is not None:
            cell = cell + ':' + xl_rowcol_to_cell(merge_to_row + row_offset, merge_to_col)
            worksheet.merge_range(cell, self.__get_casted_value(cell_value, item), format)
        elif cell_value is not None and type(cell_value) is str and cell_value.strip().startswith('='):
            worksheet.write_formula(cell, cell_value, format)
        else:
            worksheet.write(cell, self.__get_casted_value(cell_value, item), format)

        if 'validation' in item.keys():
            worksheet.data_validation(cell, item['validation'].copy())

    def __get_format(self, item):
        style = item['style'] if 'style' in item else None
        if type(style) is str:
            format = self.__formats[style]
        else:
            format = self.__workbook.add_format(style)
        return format

    def __get_merged_cells_coords(self, col):
        if ':' in col:
            columns = col.split(':')
            col = columns[0]
            merge_to_row, merge_to_col = xl_cell_to_rowcol(columns[1])
        else:
            merge_to_row, merge_to_col = None, None
        row, col = xl_cell_to_rowcol(col)
        return col, row, col, merge_to_row, merge_to_col

    def __fill_cell(self, worksheet, item):
        cell = item['cell']
        cell = self.__substitute_last_coords(cell)
        cell, row, col, merge_to_row, merge_to_col = self.__get_merged_cells_coords(cell)
        format = self.__get_format(item)
        value = item['value'] if 'value' in item else None

        self.write_value(worksheet, item, value, format, row, col, 0, merge_to_row, merge_to_col)
        return row, col

    def __fill_col(self, worksheet, item):
        col = item['col']
        col = self.__substitute_last_coords(cell)

        #cell, row, col, merge_to_row, merge_to_col = self.__get_merged_cells_coords(cell)
        if ':' in col:
            columns = col.split(':')
            col = columns[0]
            merge_to_row, merge_to_col = xl_cell_to_rowcol(columns[1])
        else:
            merge_to_row, merge_to_col = None, None
        row, col = xl_cell_to_rowcol(col)

        format = self.__get_format(item)

        if 'loop' in item.keys():
            value_list = self.__csv.values[item['loop']]
        else:
            value_list = None

        row_offset = -1
        if value_list is not None:
            idx = item['index'] if 'index' in item.keys() else None
            if idx is None:
                value = item['value'] if 'value' in item.keys() else None
            for csv_values in value_list:
                row_offset += 1
                value = csv_values[idx] if idx is not None else value
                self.write_value(worksheet, item, value, format, row, col, row_offset, merge_to_row, merge_to_col)
        else:
            for value in item['values']:
                row_offset += 1
                self.write_value(worksheet, item, value, format, row, col, row_offset, merge_to_row, merge_to_col)

        if 'spare_rows' in item:
            if 'copy_value_on_spare_row' in item and item['copy_value_on_spare_row']:
                value = item['value'] if 'value' in item.keys() else None
            else:
                value = None
            for i in range(item['spare_rows']):
                row_offset += 1
                self.write_value(worksheet, item, value, format, row, col, row_offset, merge_to_row, merge_to_col)

        return row + row_offset, col

    def __fill_row(self, worksheet, item):
        row = item['row']
        row = self.__substitute_last_coords(row)
        row, col = xl_cell_to_rowcol(row)
        format = self.__get_format(item)

        if 'loop' in item.keys():
            value_list = self.__csv.values[item['loop']]
        else:
            value_list = None

        row_offset = -1
        col_ofsset = -1
        if value_list is not None:
            for csv_values in value_list:
                col_ofsset = -1
                row_offset += 1
                for idx in item['indexes']:
                    col_ofsset += 1
                    cell = xl_rowcol_to_cell(row + row_offset, col + col_ofsset)
                    value = csv_values[idx] if idx is not None else None
                    worksheet.write(cell, self.__get_casted_value(value, item), format)
                    if 'validation' in item.keys():
                        worksheet.data_validation(cell, item['validation'].copy())
        else:
            for value in item['values']:
                col_ofsset += 1
                cell = xl_rowcol_to_cell(row, col + col_ofsset)
                worksheet.write(cell, self.__get_casted_value(value, item), format)
                if 'validation' in item.keys():
                    worksheet.data_validation(cell, item['validation'].copy())

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
