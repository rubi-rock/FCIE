import os.path
import logging


class FileLogger(object):
    def __init__(self):
        self.__filename = None

    def set_file_name(self, filename):
        self.__filename = os.path.splitext(filename)[0]+".txt"

    def __writeline(self, file, line):
        line = (line + '\n\r').encode('utf-8').decode('ISO-8859-1')
        file.write(line)

    def log(self, content):
        try:
            with open(self.__filename, 'a+', encoding='iso-8859-1') as file:
                if type(content) is str:
                    self.__writeline(file, content)
                elif type(content) is list:
                    for item in content:
                        self.__writeline(file, item)
                elif type(content) is dict:
                    for name, value in content:
                        self.__writeline(file, '{0}: {1}'.format(name, value))
                else:
                    raise TypeError("Unsuppored content type : " + str(content))
        except:
            logging.exception("Unable to write information")


AbnomaliesLogger = FileLogger()