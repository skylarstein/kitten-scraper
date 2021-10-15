from datetime import datetime
import os
import sys

class Log():
    if not sys.platform.startswith('win32'):
        GREEN     = '\033[92m'
        YELLOW    = '\033[93m'
        RED       = '\033[91m'
        CYAN      = '\033[96m'
        BOLD      = '\033[1m'
        UNDERLINE = '\033[4m'
        END       = '\033[0m'
    else:
        GREEN = YELLOW = RED = BOLD = UNDERLINE = END = ''

    @staticmethod
    def success(msg):
        print(f'{Log.GREEN}{msg}{Log.END}')

    @staticmethod
    def warn(msg):
        print(f'{Log.YELLOW}{msg}{Log.END}')

    @staticmethod
    def error(msg):
        print(f'{Log.RED}{msg}{Log.END}')

    @staticmethod
    def debug(msg):
        print(f'{Log.CYAN}{msg}{Log.END}')

class Utils():
    @staticmethod
    def utf8(strval):
        return strval.encode('utf-8').decode().strip() if sys.version_info.major >= 3 else strval.encode('utf-8').strip()

    @staticmethod
    def make_dir(fullpath):
        dirname = os.path.dirname(fullpath)
        if dirname and not os.path.exists(dirname):
            os.makedirs(dirname)

    @staticmethod
    def default_dir():
        ''' Disclaimer: I'm not trying to be wildly portable here!
        '''
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

    @staticmethod
    def levenshtein_ratio(string1, string2, strip_no_case = True):
        ''' Somewhat fuzzy string matching. Determine how closely two strings resemble each other.

            Says Wikipedia: "The Levenshtein distance between two words is the minimum number of single-character edits
            (insertions, deletions or substitutions) required to change one word into the other."
        '''
        _string1 = string1.lower().strip() if strip_no_case else string1
        _string2 = string2.lower().strip() if strip_no_case else string2

        rows = len(_string1)
        cols = len(_string2)
        distances = []

        for row in range(rows + 1):
            distances.append([row])
        for col in range(1, cols + 1):
            distances[0].append(col)

        for col in range(1, cols + 1):
            for row in range(1, rows + 1):
                if _string1[row - 1] == _string2[col - 1]:
                    distances[row].insert(col, distances[row - 1][col - 1])
                else:
                    distances[row].insert(col, min(distances[row - 1][col] + 1,      # cost of deletions
                                                   distances[row][col - 1] + 1,      # cost of insertions
                                                   distances[row - 1][col - 1] + 1)) # cost of substitutions
        distance = distances[-1][-1]
        distance_length = (rows + cols)
        return (distance_length - distance)/distance_length if distance_length else 0

    @staticmethod
    def string_to_datetime(date_str):
        ''' Attempt to convert a date string (of various formats) into datetime
        '''
        try:
            date_obj = datetime.strptime(date_str, '%d-%b-%Y')
        except ValueError:
            try:
                date_obj = datetime.strptime(date_str, '%d-%B-%Y')
            except ValueError:
                try:
                    date_obj = datetime.strptime(date_str, '%m/%d/%Y')
                except ValueError:
                    date_obj = None
        return date_obj
