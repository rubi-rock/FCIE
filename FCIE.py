import click
import os.path
import logging
import excel_builder
from os_path_helper import FileSeeker
from other_helpers import LogUtility    # keep it even if it seems unused, it set up the logging automatically
from colorama import init, Fore, Style

HELP_FILE_PARAM = 'you must provide at least one CSV file name to process. E.g.: FCIE -file MyFile.csv.\nYou can also process multiple files at once: FCIE -file MyFile1.csv -file MyFile2.csv'
HELP_FOLDER_PARAM = 'you must provide an existing path, e.g.: FCIE -folder c:/csv_files. All CSV files from this folder will be converted to excel files.'

# Colorama init
init(autoreset=True)

def print_version(ctx, param, value):
    if not value or ctx.resilient_parsing:
        return
    click.echo('Version 1.0')
    ctx.exit()


def process_file(filename):
    try:
        excel_builder.ExcelGenerator(filename)
        print(Fore.RESET + Fore.GREEN + "File processed successfuly: {0}".format(filename))
    except Exception as e:
        text = "Unable to process file: " + filename
        logging.exception(text)
        print(Fore.RESET + Fore.RED + text)
        print(Fore.RESET + Fore.LIGHTBLACK_EX + '\t\t' + str(e))

def validate_and_process_file(ctx, param, value):
    if len(value) == 0:
        return
    for file in value:
        if not os.path.exists(file):
            raise click.BadParameter('Unable to locate file: {0}'.format(file))

    for file in value:
        process_file(file)


def validate_and_process_folder(ctx, param, value):
    if value is None:
        return
    if not os.path.exists(value):
        raise click.BadParameter('Unable to locate filder: {0}'.format(value))

    file_list = FileSeeker.walk(value, ['*.csv'])
    for file in file_list:
        if file.file_name.startswith("Agence"):
            process_file(file.fullname)

@click.group()
def cli():
    """
    This tool convert CSV files extrated by Gilles Belletete's tool to build a proposition\b
    for biiling customer in the context of the new SYRA billing system of the RAMQ.\b
    \b
    This tool can work from CSV files provided one by one, from a list or from a folder.\b
    \b
    1. From file(s):\b
    ---------------\b
    You must provide at least one CSV file name to process. E.g.:\b
        FCIE -file MyFile.csv.\b
    \b
    You can also process multiple files at once:\b
        FCIE -file MyFile1.csv -file MyFile2.csv'\b
    \b
    2. from a folder:\b
    ----------------\b
    You must provide an existing path, e.g.:\b
        FCIE -folder c:/csv_files\b
    \b
    All CSV files from this folder will be converted to excel files.\b
    \b
    """
    pass


@cli.command()
@click.option('--version', is_flag=True, callback=print_version, expose_value=False, is_eager=True)
@click.option('--files', '-file', multiple=True, callback=validate_and_process_file, help=HELP_FILE_PARAM)
@click.option('-folder', type=click.Path(exists=True), callback=validate_and_process_folder, help=HELP_FOLDER_PARAM)
def convert(files, folder):
    pass


if __name__ == '__main__':
    convert()
