import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell

from csv_parser import CSVParser
from constants import EXCEL_HEADERS, ExcelBlockDef, ExcelBlock, OUI_NON

class ExcelGenerator(object):
    def __init__(self, filename):
        self.__csv = CSVParser(filename)
        self.__workbook = xlsxwriter.Workbook(filename.replace('csv', 'xlsx'))

        self.__formats = {}
        self.__init_formats()

        self.__generate_config_base()
        self.__generate_institution()
        self.__generate_user_access()
        self.__workbook.close()

    def __write_section(self, worksheet, row, col, section):
        for section in section:
            cell = xl_rowcol_to_cell(row, col)
            worksheet.write(cell, section, self.__formats['section'])
            col += 1

    def __write_headers(self, worksheet, row, col, headers):
        for header in headers:
            cell = xl_rowcol_to_cell(row, col)
            worksheet.write(cell, header, self.__formats['header'])
            col += 1

    def __write_cell(self, worksheet, row, col, value):
        cell = xl_rowcol_to_cell(row, col)
        worksheet.write(cell, value, self.__formats['cell'])

    def __init_formats(self):
        self.__formats['section'] = self.__workbook.add_format(
            {'border': 6, 'bold': True, 'font_name': 'Times New Roman'})
        self.__formats['header'] = self.__workbook.add_format(
            {'border': 6, 'bold': False, 'font_name': 'Times New Roman'})
        self.__formats['cell'] = self.__workbook.add_format({'valign': 'top', 'text_wrap': True, 'border': 1})

    def __create_worksheet(self, name):
        worksheet = self.__workbook.add_worksheet(name)
        worksheet.set_landscape()
        worksheet.set_paper(5)
        worksheet.set_margins(0.5, 0.5, 1.5, 0.8)
        worksheet.set_header('&L&G&R&F', {'image_left': 'purkinje.png'})
        worksheet.set_footer('&L&A&CPage &P / &N&R&D')
        return worksheet

    def __generate_config_base(self):
        worksheet = self.__create_worksheet("Config de base")
        self.__write_section(worksheet, 0, 0, EXCEL_HEADERS[ExcelBlock.customer][ExcelBlockDef.section])
        self.__write_headers(worksheet, 1, 0, EXCEL_HEADERS[ExcelBlock.customer][ExcelBlockDef.headers])
        self.__write_section(worksheet, 4, 0, EXCEL_HEADERS[ExcelBlock.user][ExcelBlockDef.section])
        self.__write_headers(worksheet, 5, 0, EXCEL_HEADERS[ExcelBlock.user][ExcelBlockDef.headers])
        self.__write_section(worksheet, 4, 5, EXCEL_HEADERS[ExcelBlock.md][ExcelBlockDef.section])
        self.__write_headers(worksheet, 5, 5, EXCEL_HEADERS[ExcelBlock.md][ExcelBlockDef.headers])
        worksheet.set_column('A:A', 25)
        worksheet.set_column('B:M', 15)
        worksheet.set_column('J:J', 40)
        worksheet.set_column('N:N', 20)
        worksheet.set_column('E:E', 5)  #spacer
        #Customers
        customer = self.__csv.values[ExcelBlock.customer][0]
        self.__write_cell(worksheet, 2, 0, None)    # Nom du client / agence
        self.__write_cell(worksheet, 2, 1, int(customer[0]))   # Numéro agence
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
            row += 1
        # MD
        mds = self.__csv.values[ExcelBlock.md]
        row = 6
        for md in mds:
            self.__write_cell(worksheet, row, 5, int(md[0]))  # Numero de pratique
            self.__write_cell(worksheet, row, 6, int(md[1]) if md[1] != '' else None)  # Numero de groupe
            self.__write_cell(worksheet, row, 7, md[2])  # Nom
            self.__write_cell(worksheet, row, 8, md[3])  # Prénon
            self.__write_cell(worksheet, row, 9, md[4])  # Specialité
            self.__write_cell(worksheet, row, 10, OUI_NON[md[5]])  # RMX (oui/non)
            self.__write_cell(worksheet, row, 11, None)  # Inc (oui/non)
            self.__write_cell(worksheet, row, 12, None)  # Nom compagnie
            self.__write_cell(worksheet, row, 13, None)  # Date fin année fiscalle inc
            row += 1

    def __generate_institution(self):
        worksheet = self.__create_worksheet("Établissement")
        self.__write_section(worksheet, 0, 0, EXCEL_HEADERS[ExcelBlock.institution][ExcelBlockDef.section])
        self.__write_headers(worksheet, 1, 0, EXCEL_HEADERS[ExcelBlock.institution][ExcelBlockDef.headers])

    def __generate_user_access(self):
        worksheet = self.__create_worksheet("Accès utilisateurs")
        self.__write_section(worksheet, 0, 0, EXCEL_HEADERS[ExcelBlock.user_access][ExcelBlockDef.section])
        self.__write_headers(worksheet, 1, 0, EXCEL_HEADERS[ExcelBlock.user_access][ExcelBlockDef.headers])
