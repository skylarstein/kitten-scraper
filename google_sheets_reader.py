import pygsheets
from kitten_utils import *

class GoogleSheetsReader(object):
    def load_mentors_spreadsheet(self, sheets_key):
        ''' Load the feline foster spreadsheet
        '''
        self.mentor_sheets = []
        try:
            print_success('Loading mentors spreadsheet {}...'.format(sheets_key))
            client = pygsheets.authorize(outh_file='client_secret.json')
            spreadsheet = client.open_by_key(sheets_key)

            config_yaml = spreadsheet.worksheet_by_title("Config")[1][0]

            for worksheet in spreadsheet.worksheets():
                if worksheet.title.lower() not in ['contact info', 'config', 'updates', 'announcements', 'resources']:
                    self.mentor_sheets.append(worksheet)
        except Exception as e:
            print_err('ERROR: Unable to load Feline Foster spreadsheet!\r\n{}, {}'.format(str(e), repr(e)))
            return None

        print('Loaded {} mentors from spreadsheet'.format(len(self.mentor_sheets)))
        return config_yaml

    def find_matches_in_feline_foster_spreadsheet(self, match_strings):
        ''' Find mentor worksheets that match any string in match_strings. Not very sophisticated right now, I'm simply
            searching for a match anywhere in each mentor sheet.
        '''
        match_strings = [utf8(s).lower() for s in match_strings if s]
        matching_mentors = set()

        for sheet in self.mentor_sheets:
            all_values = sheet.get_all_values(include_tailing_empty=False, include_tailing_empty_rows=False)
            flattend = [utf8(item).lower() for sublist in all_values for item in sublist]
            if len([item for item in flattend if any(match in item for match in match_strings)]):
                matching_mentors.add(utf8(sheet.title))

        return matching_mentors
