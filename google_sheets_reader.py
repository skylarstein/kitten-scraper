import pygsheets
from kitten_utils import *
from sheet_reader_base import SheetReaderBase

class GoogleSheetsReader(SheetReaderBase):
    def load_mentors_spreadsheet(self, auth):
        ''' Load the feline foster spreadsheet
        '''
        self.mentor_sheets = []
        self.flattend_sheet_values = {}

        try:
            print_success('Loading mentors spreadsheet {}...'.format(auth['google_spreadsheet_key']))
            client = pygsheets.authorize(auth['google_client_secret'])
            spreadsheet = client.open_by_key(auth['google_spreadsheet_key'])

            config_yaml = spreadsheet.worksheet_by_title("Config")[1][0]

            for worksheet in spreadsheet.worksheets():
                if worksheet.title.lower() not in ['contact info', 'config', 'updates', 'announcements', 'resources', 'calendar']:
                    self.mentor_sheets.append(worksheet)
                    all_values = worksheet.get_all_values(include_tailing_empty = False, include_tailing_empty_rows = False)
                    self.flattend_sheet_values[utf8(worksheet.title)] = [utf8(item).lower() for sublist in all_values for item in sublist]

        except Exception as e:
            print_err('ERROR: Unable to load Feline Foster spreadsheet!\r\n{}, {}'.format(str(e), repr(e)))
            return None

        print('Loaded {} mentors from spreadsheet'.format(len(self.mentor_sheets)))
        return config_yaml

    def find_matching_mentors(self, match_strings):
        ''' Find mentor worksheets that match any string in match_strings. Not very sophisticated right now, I'm simply
            searching for a match anywhere in each mentor sheet.
        '''
        match_strings = [utf8(s).lower() for s in match_strings if s]
        matching_mentors = set()

        for sheet_name in self.flattend_sheet_values:
            if len([item for item in self.flattend_sheet_values[sheet_name] if any(match in item for match in match_strings)]):
                matching_mentors.add(sheet_name)

        return matching_mentors

    def get_current_mentees(self):
        ''' Return the current mentees assigned to each mentor
        '''
        current_mentees = []
        for worksheet in self.mentor_sheets:
            if worksheet.title.lower() == 'retired mentor':
                continue
            print('Loading current mentees for {}... '.format(worksheet.title), end='')
            mentees = []

            # It's much faster to grab a whole block of cells at once vs iterating through many API calls
            #
            max_search_rows = 50
            cells = worksheet.range('A1:G{}'.format(max_search_rows), returnas='cells')

            name_col_id = self._find_column_by_name(cells, 'Name')
            pid_col_id = self._find_column_by_name(cells, 'ID')

            search_failed = False
            for i in range(1, max_search_rows):
                if i == max_search_rows - 1:
                    search_failed = True
                    print_err('Unable to determine current mentees for mentor {}'.format(worksheet.title))
                    mentees = []
                    break

                elif cells[i][0].value.lower().strip() == 'completed mentees without kittens':
                    break # We've reach the end of "active mentee" rows

                elif cells[i][name_col_id].value and cells[i][pid_col_id].value:
                    mentee_name = cells[i][name_col_id].value
                    pid = int(cells[i][pid_col_id].value)
                    if not [mentee for mentee in mentees if mentee['pid'] == pid]: # ignore duplicate mentees
                        mentees.append({'name' : mentee_name, 'pid' : pid})

            if not search_failed:
                print('found {}'.format(len(mentees)))

            current_mentees.append({ 'mentor' : worksheet.title, 'mentees' : mentees})

        return current_mentees

    def _find_column_by_name(self, cells, name):
        for n in range(0, len(cells[0])):
            if cells[0][n].value == name:
                return n
        return 0
