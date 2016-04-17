import re
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol
import ast
import logging

from csv_parser import CSVParser
from constants import OUI_NON, OS_LIST


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
            elif name == 'page_view' and value:
                worksheet.set_page_view()

    def __instanciate_worksheet(self, ws_definition):
        try:
            worksheet = self.__workbook.add_worksheet(ws_definition['name'])
            self.__reformat_worksheet(worksheet, ws_definition['format'])
            worksheet.set_margins(0.5, 0.5, 1.5, 0.8)
            if ws_definition['header'] is not None and ws_definition['header']['format'] is not None:
                worksheet.set_header(ws_definition['header']['format'], ws_definition['header']['options'])
            if ws_definition['footer'] is not None and ws_definition['footer']['format'] is not None:
                worksheet.set_footer(ws_definition['footer']['format'], ws_definition['footer']['options'] if 'option' in ws_definition['footer'].keys() else None)
            return worksheet
        except:
            logging.exception('failed to create tab: ' + ws_definition['name'])

    def __add_page_break(self, worksheet, item):
        row = self.__substitute_variables(item['break'])
        self.__variables['breaks'].append(row)

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
                elif 'break' in params.keys():
                    ws_definition['content'].append(params)
                elif 'table' in params.keys():
                    ws_definition['content'].append(params)
                elif 'col' in params.keys():
                    ws_definition['content'].append(params)
                elif 'row' in params.keys():
                    ws_definition['content'].append(params)
                elif 'vspace' in params.keys():
                    ws_definition['content'].append(params)
                elif 'hspace' in params.keys():
                    ws_definition['content'].append(params)

        return self.__process_ws_definition(ws_definition)

    def __process_ws_definition(self, ws_definition):
        try:
            worksheet =  self.__instanciate_worksheet(ws_definition)
            self.__variables.clear()
            self.__variables['breaks'] = []
            self.__variables['last_row'], self.__variables['last_column'] = 0, 0
            for item in ws_definition['content']:
                if 'cell' in item.keys():
                    lr, lc = self.__fill_cell(worksheet, item)
                elif 'col' in item.keys():
                    lr, lc = self.__fill_col(worksheet, item)
                elif 'row' in item.keys():
                    lr, lc = self.__fill_row(worksheet, item)
                elif 'table' in item.keys():
                    self.__add_table(worksheet, item)
                elif 'break' in item.keys():
                    self.__add_page_break(worksheet, item)
                elif 'vspace' in item.keys():
                    self.__variables['last_row'] = self.__variables['last_row'] + item['vspace']
                elif 'hspace' in item.keys():
                    self.__variables['last_column'] = self.__variables['last_column'] + item['hspace']

                if 'remember_last_row' in item:
                    self.__variables['last_row'] = lr + 1
                if 'remember_last_column' in item:
                    self.__variables['last_column'] = lc + 1

            if 'breaks' in self.__variables:
                worksheet.set_h_pagebreaks(self.__variables['breaks'])

            return worksheet
        except:
            logging.exception('Unable to process tab definition: ' + str(ws_definition))
            return None

    def __add_table(self, worksheet, item):
        coords = self.__substitute_variables(item['table'])
        worksheet.add_table(coords, {'header_row': False, 'autofilter': False, 'banded_rows': False, 'banded_columns': False, 'first_column': False, 'last_column': False})

    def __prepare_eval_expression(self, py_code):
        if py_code is None or type(py_code) is not str :
            return
        for var_name in self.__variables.keys():
            if var_name in py_code:
                py_code = py_code.replace( var_name, 'self.variables["' + var_name + '"]')
        return py_code

    def __substitute_variables(self, cell_value):
        if cell_value is None or type(cell_value) is not str :
            return cell_value
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

    def write_value(self, worksheet, item, cell_value, format, row, col, row_offset, col_offset, merge_to_row, merge_to_col):
        cell = xl_rowcol_to_cell(row + row_offset, col + col_offset)
        self.__variables['current_row'] = row + row_offset + 1
        self.__variables['current_col'] = col + col_offset + 1
        cell_value = self.__substitute_variables(cell_value)

        if 'height' in item:
            worksheet.set_row(row, item['height'])

        if merge_to_col is not None:
            cell = cell + ':' + xl_rowcol_to_cell(merge_to_row + row_offset, merge_to_col)
            worksheet.merge_range(cell, self.__get_casted_value(cell_value, item), format)
        elif cell_value is not None and type(cell_value) is str and cell_value.strip().startswith('='):
            worksheet.write_formula(cell, cell_value, format)
        else:
            if 'textbox' in item:
                worksheet.insert_textbox(cell, self.__get_casted_value(cell_value, item), item['textbox'])
            else:
                worksheet.write(cell, self.__get_casted_value(cell_value, item), format)

        if 'validation' in item.keys():
            worksheet.data_validation(cell, item['validation'].copy())

    def __get_format(self, item):
        try:
            style = item['style'] if 'style' in item else None
            if style is None:
                return None

            if type(style) is str:
                format = self.__formats[style]
            else:
                format = self.__workbook.add_format(style)
            return format
        except:
            logging.exception('Unable to find style: ' + style)
            return None

    def __get_merged_cells_coords(self, col):
        if ':' in col:
            columns = col.split(':')
            col = columns[0]
            merge_to_row, merge_to_col = xl_cell_to_rowcol(columns[1])
        else:
            merge_to_row, merge_to_col = None, None
        row, col = xl_cell_to_rowcol(col)
        return col, row, col, merge_to_row, merge_to_col

    def __save_variable(self, item):
        if not 'save' in item.keys():
            return
        var = item['save'].split('=', 1)
        py_code = self.__prepare_eval_expression(var[1])
        self.variables[var[0]] = eval(py_code)

    def __fill_cell(self, worksheet, item):
        cell = item['cell']
        cell = self.__substitute_last_coords(cell)
        cell, row, col, merge_to_row, merge_to_col = self.__get_merged_cells_coords(cell)
        format = self.__get_format(item)
        value = item['value'] if 'value' in item else None

        self.__save_variable(item)
        self.write_value(worksheet, item, value, format, row, col, 0, 0, merge_to_row, merge_to_col)
        return row, col

    def __fill_col(self, worksheet, item):
        def write_col_value(row_offset):
            row_offset += 1
            self.write_value(worksheet, item, value, format, row, col, row_offset, 0, merge_to_row, merge_to_col)
            if row_offset == 0:
                self.__save_variable(item)
            return row_offset

        cell = item['col']
        cell = self.__substitute_last_coords(cell)
        cell, row, col, merge_to_row, merge_to_col = self.__get_merged_cells_coords(cell)
        format = self.__get_format(item)

        value_list = self.__csv.values[item['loop']] if 'loop' in item.keys() else None

        row_offset = -1
        if value_list is not None:
            idx = item['index'] if 'index' in item.keys() else None
            if idx is None:
                value = item['value'] if 'value' in item.keys() else None
            for csv_values in value_list:
                value = csv_values[idx] if idx is not None else value
                row_offset = write_col_value(row_offset)
        else:
            for value in item['values']:
                row_offset = write_col_value(row_offset)

        if 'spare_rows' in item:
            if 'copy_value_on_spare_row' in item and item['copy_value_on_spare_row']:
                value = item['value'] if 'value' in item.keys() else None
            else:
                value = None
            for i in range(item['spare_rows']):
                row_offset = write_col_value(row_offset)

        return row + row_offset, col

    def __fill_row(self, worksheet, item):
        def write_row_value():
            self.write_value(worksheet, item, value, format, row, col, row_offset if row_offset >= 0 else 0, col_offset if col_offset >= -1 else 0, merge_to_row, merge_to_col)
            if row_offset == 0 or col_offset == 0:
                self.__save_variable(item)

        cell = item['row']
        cell = self.__substitute_last_coords(cell)
        cell, row, col, merge_to_row, merge_to_col = self.__get_merged_cells_coords(cell)
        format = self.__get_format(item)

        value_list = self.__csv.values[item['loop']] if 'loop' in item.keys() else None

        row_offset = -1
        col_offset = -1
        if value_list is not None:
            for csv_values in value_list:
                col_offset = -1
                row_offset += 1
                for idx in item['indexes']:
                    col_offset += 1
                    value = csv_values[idx] if idx is not None else None
                    write_row_value()
        else:
            for value in item['values']:
                col_offset += 1
                write_row_value()

        return row + row_offset, col + col_offset

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


