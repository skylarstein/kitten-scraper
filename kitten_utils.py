import sys

class ConsoleFormat(object):
    GREEN     = '\033[92m'
    YELLOW    = '\033[93m'
    RED       = '\033[91m'
    BOLD      = '\033[1m'
    UNDERLINE = '\033[4m'
    END       = '\033[0m'

def print_success(msg):
    print('{}{}{}'.format(ConsoleFormat.GREEN, msg, ConsoleFormat.END))

def print_warn(msg):
    print('{}{}{}'.format(ConsoleFormat.YELLOW, msg, ConsoleFormat.END))

def print_err(msg):
    print('{}{}{}'.format(ConsoleFormat.RED, msg, ConsoleFormat.END))

def utf8(strval):
    return strval.encode('utf-8').decode().strip() if sys.version_info.major >= 3 else strval.encode('utf-8').strip()
