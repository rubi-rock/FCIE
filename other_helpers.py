import os.path
import logging
import datetime
import time
from enum import Enum
from collections import OrderedDict
import cProfile

from os_path_helper import create_folder_if_doest_not_exists

logger = logging.getLogger()

'''
    Use of the CProfile library. To perform a profiling of the code you just have to:
        . install CProfile (pip install cProfile)
        . write a script that call the method to profile:
            >>> import cProfile
            >>> cProfile.run('2 + 2')
                2 function calls in 0.000 seconds

    Another way to use it is : on the all application:
        . from a shel call the profiler with your script in parameter:
            $ cprofilev  main.py
        . it creates a file profile_output that is unreadable ;) Don't panic, just open the URL to be able to read it:
            http://127.0.0.1:4000/
        . don't need to wait the end of the execution of your script to see result, you can check the
          web page as soon as the profiler is started. You can refresh the page for an update during your script
          execution
        . The tool does not close to let you access the results web page. So you have to stop the profiler by pressing
          Ctrl+C to terminate it.

    Another way if you have KCacheGrind installed:
        . start the profiler bu specifying the output file (e.g.: myscript.prof) instead of letting cProfile name it
          profile_outout:
            $ python -m cProfile -o myscript.cprof myscript.py
        . start KCacheGrind with your .prof file to see the graphical results presented on screen
            $ pyprof2calltree -k -i myscript.cprof

    OR if you like graphics, install pycallgraph and dot then:
        . install dot with brew:
            brew install graphviz
        . install pycallgraph:
            pip3 install pycallgraph
        . start the profiler:
            pycallgraph graphviz ./myscript.py
        . check the .png file generate
'''


# cProfiler class
def do_cprofile(func):
    def profiled_func(*args, **kwargs):
        profile = cProfile.Profile()
        try:
            profile.enable()
            result = func(*args, **kwargs)
            profile.disable()
            return result
        finally:
            profile.print_stats()

    return profiled_func


# Test procedure to validate the CProfiler is working
def get_number():
    for x in range(5000000):
        yield x


# Log Utility
class LogUtility(object):
    ROOT_PATH = os.path.dirname(os.path.abspath(__file__))
    CONFIG_PATH = os.path.join(ROOT_PATH, "config")
    LOG_PATH = os.path.join(ROOT_PATH, 'log')

    __console_logger = logging.StreamHandler()
    __file_logger = logging.StreamHandler()

    @property
    def console_logger(self):
        return self.__console_logger

    @property
    def file_logger(self):
        return self.__file_logger

    def __init__(self):
        # Creates required folders if they don't exist yet
        create_folder_if_doest_not_exists(self.LOG_PATH)
        create_folder_if_doest_not_exists(self.CONFIG_PATH)

        # Very important, the root logger must let go all logs to the handlers no matter the message log level because
        # the log level must be manage at the handler level and must not be filtered out by the root logger.
        logger.setLevel(logging.NOTSET)

        # Log format (csv style)
        formatter = logging.Formatter('%(asctime)s, %(levelname)s, "%(message)s"', "%Y/%m/%d, %H:%M:%S")

        # create console handler and set level to info
        self.__console_logger.setLevel(logging.INFO)
        self.__console_logger.setFormatter(formatter)
        logger.addHandler(self.__console_logger)

        # create error file handler and set level to error
        self.__file_logger = logging.FileHandler(os.path.join(self.LOG_PATH, "error.log"), "a", encoding=None,
                                                 delay="false")
        self.__file_logger.setLevel(logging.ERROR)
        self.__file_logger.setFormatter(formatter)
        logger.addHandler(self.__file_logger)


'''
    Singleton that helps to tune logging into console and a file as we need.
'''
logging_helper = LogUtility()


# Process timer: exclude sleping time, other process time...
class ProcessTimer(object):
    def __init__(self):
        self.__start_time = time.process_time()

    @property
    def time(self):
        return time.process_time() - self.__start_time

    @property
    def time_to_str(self):
        return str(self.time)


# Elapsed Timer : just count the time like your watch between to moments
class ElapseTimer(object):
    def __init__(self):
        self.__start_time = time.perf_counter()

    @property
    def time(self):
        return time.perf_counter() - self.__start_time

    @property
    def time_to_str(self):
        dt = datetime.timedelta(seconds=self.time)
        return str(dt)


# HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!!
# Inerited class from a dict that allows to behave as:
#   . a dictionary : key - value
#   . an enum : does not require to write long and complex syntex of the enum. I.e.:
#           myenum.red
#     instead of
#           myenum.red.key
#   . It does not requires to write down a property manually for each of the values of the enum and maintain it when
#     it's not a native Enum class of Python
#   . also support dict serialization (to print) and other native behaviors like keys(), values() for serialization (csv)
#
#   Warning: it's not read only, so don't add values because it wont generate a property, it's not implemented
#
#   Warning: no code completion on this "enum", you must know your keywords
#
class DictEnum(OrderedDict):
    # Initialize the dict from another one containing the enum list Names & values, plus publish all keys as properties
    # directly available from this object
    def __init__(self, enum_list, *args, **kwds):
        # load the list
        super().__init__(*args, **kwds)
        if type(enum_list) is dict or type(enum_list) is OrderedDict:
            self.update(enum_list)
        elif type(enum_list) is tuple or type(enum_list) is list:
            for name in enum_list:
                self[name] = name
        else:
            raise TypeError()
        # create the properties
        for key, value in self.items():
            object.__setattr__(self, key, value)


# HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!!
#
# Same idea than the DictEnum but simpler and keep original order for the enumerated items which is useful to
# dump in a CSV file
#
class ListEnum(list):
    # Initialize the dict from another one containing the enum list Names & values, plus publish all keys as properties
    # directly available from this object
    def __init__(self, enum_list):
        # keep a trace of the list for the 'as_list' operator
        super().__init__()
        self.__enum_list = enum_list
        # load the list
        for name in enum_list:
            self.append(name)
        # create the properties
        [object.__setattr__(self, name.replace(' ', '_'), name) for name in enum_list]

    @property
    def as_list(self):
        return self.__enum_list

# HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!! HACK!! HACK!!! HACK!!!
#
# Version of the Enum more convenient, still you need to add .value to an enum to be able to us it... too cumbersome
#
class EnumEx(Enum):
    def __str__(self):
        return self.value

    @property
    def names(self):
        cls = self.__class__
        return [enum_item.name for enum_item in cls]

    @property
    def values(self):
        cls = self.__class__
        return [enum_item.value for enum_item in cls]

    @property
    def dictionary(self):
        return dict(zip(self.names, self.values))
