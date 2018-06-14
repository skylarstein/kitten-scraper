class ConsoleFormat(object):
    SUCCESS   = '\033[92m'
    WARNING   = '\033[93m'
    ERROR     = '\033[91m'
    BOLD      = '\033[1m'
    UNDERLINE = '\033[4m'
    END       = '\033[0m'

def print_success(msg):
    print('{}{}{}'.format(ConsoleFormat.SUCCESS, msg, ConsoleFormat.END))

def print_warn(msg):
    print('{}{}{}'.format(ConsoleFormat.WARNING, msg, ConsoleFormat.END))

def print_err(msg):
    print('{}{}{}'.format(ConsoleFormat.ERROR, msg, ConsoleFormat.END))

def utf8(strval):
    return strval.encode('utf-8').strip()