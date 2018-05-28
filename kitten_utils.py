class Colors:
    SUCCESS   = '\033[92m'
    WARNING   = '\033[93m'
    ERROR     = '\033[91m'
    BOLD      = '\033[1m'
    UNDERLINE = '\033[4m'
    END       = '\033[0m'

def print_success(msg):
    print('{}{}{}'.format(Colors.SUCCESS, msg, Colors.END))

def print_warn(msg):
    print('{}{}{}'.format(Colors.WARNING, msg, Colors.END))

def print_err(msg):
    print('{}{}{}'.format(Colors.ERROR, msg, Colors.END))
