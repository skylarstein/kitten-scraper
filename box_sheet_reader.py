import os
import xlrd
from boxsdk import Client, JWTAuth
from kitten_utils import *
from sheet_reader_base import SheetReaderBase

class BoxSheetReader(SheetReaderBase):
    def __init__(self):
        super().__init__()
        self._mentor_sheets = []
        self._flattend_sheet_values = {}

    def load_mentors_spreadsheet(self, auth):
        ''' Load the feline foster spreadsheet
        '''
        try:
            print_success('Loading mentors spreadsheet from Box ({})...'.format(auth['box_file_id']))

            jwt_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), auth['box_jwt'])
            client = Client(JWTAuth.from_settings_file(jwt_path))
            box_file = client.as_user(client.user(user_id = auth['box_user_id'])).file(file_id=auth['box_file_id']).get()
            xlxs_workbook = xlrd.open_workbook(file_contents=box_file.content())

            config_yaml = xlxs_workbook.sheet_by_name(self._config_sheet_name).row_values(1)[0]

            for sheet_name in xlxs_workbook.sheet_names():
                if not self._is_reserved_sheet(sheet_name):
                    sheet = xlxs_workbook.sheet_by_name(sheet_name)
                    self._mentor_sheets.append(sheet)
                    all_values = [sheet.row_values(i) for i in range(1, sheet.nrows)]
                    self._flattend_sheet_values[utf8(sheet_name)] = [utf8(str(item)).lower() for sublist in all_values for item in sublist]

        except Exception as e:
            print_err('ERROR: Unable to load Feline Foster spreadsheet!\r\n{}, {}'.format(str(e), repr(e)))
            return None

        print('Loaded {} mentors from \"{}\"'.format(len(self._mentor_sheets), box_file['name']))
        return config_yaml

    def find_matching_mentors(self, match_strings):
        ''' Find mentor worksheets that match any string in match_strings. Not very sophisticated right now, I'm simply
            searching for a match anywhere in each mentor sheet.
        '''
        match_strings = [utf8(s).lower() for s in match_strings if s]
        matching_mentors = set()

        for sheet_name in self._flattend_sheet_values:
            if len([item for item in self._flattend_sheet_values[sheet_name] if any(match in item for match in match_strings)]):
                matching_mentors.add(sheet_name)

        return matching_mentors

    def get_current_mentees(self):
        ''' Return the current mentees assigned to each mentor
        '''
        current_mentees = []
        return current_mentees

    def _find_column_by_name(self, cells, name):
        for n in range(0, len(cells[0])):
            if cells[0][n].value == name:
                return n
        return 0
