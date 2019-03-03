import pygsheets
from kitten_utils import *

class GoogleSheetsReader(object):
    def load_mentors_spreadsheet(self, sheets_key):
        ''' Load the feline foster spreadsheet
        '''
        self._mentor_data = []
        try:
            print_success('Loading mentors spreadsheet {}...'.format(sheets_key))
            gc = pygsheets.authorize(outh_file='client_secret.json')
            spreadsheet = gc.open_by_key(sheets_key)
            worksheets = spreadsheet.worksheets()

            config_sheet = spreadsheet.worksheet_by_title("Config")
            config_yaml = config_sheet[1][0]

            for worksheet in worksheets:
                if worksheet.title.lower() in ['resources', 'config', 'updates', 'announcements']:
                    continue
                row_data = []
                for row in worksheet:
                    row_data.append(','.join(utf8(val).lower() for val in row)) # save whole row as CSV for now

                self._mentor_data.append({worksheet.title : row_data})
        except Exception as e:
            print_err('ERROR: Unable to load Feline Foster spreadsheet!\r\n{}, {}'.format(str(e), repr(e)))
            return None

        print('Loaded {} mentors from spreadsheet'.format(len(self._mentor_data)))
        return config_yaml

    def find_matches_in_feline_foster_spreadsheet(self, match_strings):
        ''' Find mentor worksheets that match any string in match_strings. Not very sophisticated.
        '''
        match_strings = [utf8(s).lower() for s in match_strings if s]
        matching_sheets = set()
        for sheet in self._mentor_data:
            sheet_name, sheet_rows = list(sheet.items())[0]
            for match_string in match_strings:
                if next((row for row in sheet_rows if match_string in row), None):
                    matching_sheets.add(utf8(sheet_name))
                    break

        return matching_sheets
