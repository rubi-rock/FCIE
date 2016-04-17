import click
import excel_builder
from other_helpers import LogUtility    # keep it even if it seems unused, it set up the logging automatically



if __name__ == '__main__':
    eg = excel_builder.ExcelGenerator('test2.csv')