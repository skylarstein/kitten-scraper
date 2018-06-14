import pygsheets
from kitten_utils import *

class GoogleSheetsReader(object):
    def load_mentors_spreadsheet(self, sheets_key):
        ''' Load the feline foster spreadsheet
        '''
        self._sheet_data = []
        try:
            print_success('Loading mentors spreadsheet {}...'.format(sheets_key))
            gc = pygsheets.authorize(outh_file='client_secret.json')
            spreadsheet = gc.open_by_key(sheets_key)
            worksheets = spreadsheet.worksheets()

            for n in range(2, len(worksheets)): # ignore first two sheets
                row_data = []
                worksheet = spreadsheet.worksheet('index', n)
                for row in worksheet:
                    row_data.append(','.join(utf8(val).lower() for val in row)) # save whole row as CSV for now

                self._sheet_data.append({worksheet.title : row_data})
        except Exception, e:
            print_err('ERROR: Unable to load Feline Foster spreadsheet!\r\n{}'.format(str(e)))
            return False

        print('Loaded {} worksheets from mentors spreadsheet'.format(len(self._sheet_data)))
        return True

    def find_matches_in_feline_foster_spreadsheet(self, match_strings):
        ''' Find mentor worksheets that match any string in match_strings. Not very sophisticated.
        '''
        match_strings = [utf8(s).lower() for s in match_strings if s]
        matching_sheets = set()
        for sheet in self._sheet_data:
            sheet_name, sheet_rows = sheet.items()[0]
            for match_string in match_strings:
                if next((row for row in sheet_rows if match_string in row), None):
                    matching_sheets.add(utf8(sheet_name))
                    break

        return matching_sheets
