import os
import sys

class ConsoleFormat(object):
    if not sys.platform.startswith('win32'):
        GREEN     = '\033[92m'
        YELLOW    = '\033[93m'
        RED       = '\033[91m'
        BOLD      = '\033[1m'
        UNDERLINE = '\033[4m'
        END       = '\033[0m'
    else:
        GREEN = YELLOW = RED = BOLD = UNDERLINE = END = ''

def print_success(msg):
    print('{}{}{}'.format(ConsoleFormat.GREEN, msg, ConsoleFormat.END))

def print_warn(msg):
    print('{}{}{}'.format(ConsoleFormat.YELLOW, msg, ConsoleFormat.END))

def print_err(msg):
    print('{}{}{}'.format(ConsoleFormat.RED, msg, ConsoleFormat.END))

def utf8(strval):
    return strval.encode('utf-8').decode().strip() if sys.version_info.major >= 3 else strval.encode('utf-8').strip()

def default_dir():
    # Disclaimer: I'm not trying to be wildly portable here!
    #
    if sys.platform == 'darwin' or sys.platform.startswith('linux'):
        return os.path.expanduser('~/Desktop')
    elif sys.platform.startswith('win32'):
        desktop_onedrive = os.path.join(os.environ['USERPROFILE'], 'OneDrive')
        if os.path.exists(desktop_onedrive):
            return desktop_onedrive
        else:
            return os.path.join(os.environ['USERPROFILE'], 'Desktop')
    else:
        return os.path.dirname(os.path.realpath(__file__))
