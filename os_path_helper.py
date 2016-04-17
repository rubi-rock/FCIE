import logging
import os
import fnmatch
from datetime import datetime
import inspect
import sys

import subprocess


class FileInfo(dict):
    # def initialize(self, root, filename):
    def __init__(self, root, filename):
        super(FileInfo, self).__init__()

        absolute_filename = os.path.join(root, filename)
        self['file_name'] = os.path.basename(absolute_filename)
        self['path'] = os.path.dirname(absolute_filename)
        self['date'] = datetime.fromtimestamp(os.stat(self.fullname).st_mtime)

    def __str__(self):
        return "Path: {0}, File Name: {1}, Changed Date: {2}".format(self.path, self.file_name, self.date)

    @property
    def file_name(self):
        return self['file_name']

    @property
    def path(self):
        return self['path']

    @property
    def date(self):
        return self['date']

    @property
    def fullname(self):
        return os.path.join(self.path, self.file_name)


class FileSeeker(object):
    __default_pattern = "*.*"

    @staticmethod
    def walk(root_path, extensions=None):
        file_list = []
        for root, dirs, files in os.walk(root_path):
            for ext in extensions:
                for filename in fnmatch.filter(files, ext):
                    lfi = FileInfo(root, filename)
                    file_list.append(lfi)
        return file_list

    '''
    # Example of method that can be used. It works with functions, methods and static methods.
    # The parameter is a LogFileInfo object, and the function must return true if the log
    # file must be in the final list
    def filterLogs(lfi):
        return "Scheduler" in lfi.FileName
    '''

    @staticmethod
    def walk_and_filter_in(root_path, extensions=None, filter_callback=None):
        file_list = FileSeeker.walk(root_path, extensions)
        if inspect.ismethod(filter_callback) or inspect.isfunction(filter_callback):
            result = [x for x in file_list if filter_callback(x)]
        else:
            logging.debug("Invalid filter callback. The filtering must be a function, a method or a static method")
            result = file_list
        return result


def create_folder_if_doest_not_exists(folder_name):
    if not os.path.exists(folder_name):
        try:
            os.mkdir(folder_name)
        except Exception as e:
            logging.exception("Unable to find or create the folder " + folder_name, e)
            sys.exit()

def generate_file_name(prefix='', extension = '.tmp', path = os.path.curdir):
    filename = '%s' % datetime.now().strftime('%Y%m%d-%H%M%S')
    filename = prefix + filename if prefix is not None and prefix != '' else filename
    filename = filename + extension if extension is not None and extension != '' else filename
    return os.path.join(path, filename)

def start_file(filename):
    if sys.platform == "win32":
        os.startfile(filename)
    elif sys.platform == "darwin":
        subprocess.call(["open", filename])