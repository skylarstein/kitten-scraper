import pygsheets
from kitten_utils import *
from sheet_reader_base import SheetReaderBase

class GoogleSheetReader(SheetReaderBase):
    def __init__(self):
        super().__init__()

    def load_mentors_spreadsheet(self, auth):
        ''' Load the feline foster spreadsheet
        '''
        try:
            print_success('Loading mentors spreadsheet from Google Sheets (id = {})...'.format(auth['google_spreadsheet_key']))

            client = pygsheets.authorize(auth['google_client_secret'])
            spreadsheet = client.open_by_key(auth['google_spreadsheet_key'])

            config_yaml = spreadsheet.worksheet_by_title(self._config_sheet_name)[2][1]

            for worksheet in spreadsheet.worksheets():
                if not self._is_reserved_sheet(worksheet.title):
                    self._mentor_sheets.append(worksheet)
                    all_values = worksheet.get_all_values(include_tailing_empty = False, include_tailing_empty_rows = False)
                    self._flattend_sheet_values[utf8(worksheet.title)] = [utf8(item).lower() for sublist in all_values for item in sublist]

        except Exception as e:
            print_err('ERROR: Unable to load Feline Foster spreadsheet!\r\n{}, {}'.format(str(e), repr(e)))
            return None

        print('Loaded {} mentors from \"{}\"'.format(len(self._mentor_sheets), spreadsheet.title))
        return config_yaml

    def get_current_mentees(self):
        ''' Return the current mentees assigned to each mentor
        '''
        current_mentees = []
        for worksheet in self._mentor_sheets:
            if worksheet.title.lower() == 'retired mentor':
                continue
            print('Loading current mentees for {}... '.format(worksheet.title), end='')

            # It's much faster to grab a whole block of cells at once vs iterating through many API calls
            #
            max_search_rows = min(50, worksheet.rows)
            cells = worksheet.range('A1:G{}'.format(max_search_rows), returnas='cells')

            name_col_id = self._find_column_by_name(cells, 'Name')
            pid_col_id = self._find_column_by_name(cells, 'ID')

            mentees = []
            search_failed = False
            for i in range(1, max_search_rows):
                if i == max_search_rows - 1:
                    search_failed = True
                    print_err('Unable to determine current mentees for mentor {}'.format(worksheet.title))
                    mentees = []
                    break

                elif str(cells[i][0].value).lower().find('completed mentees') >= 0:
                    break # We've reach the end of the "active mentee" rows

                elif cells[i][name_col_id].value and cells[i][pid_col_id].value:
                    mentee_name = cells[i][name_col_id].value
                    pid = int(cells[i][pid_col_id].value)
                    if not [mentee for mentee in mentees if mentee['pid'] == pid]: # ignore duplicate mentees
                        mentees.append({'name' : mentee_name, 'pid' : pid})

            if not search_failed:
                print('found {}'.format(len(mentees)))

            current_mentees.append({ 'mentor' : worksheet.title, 'mentees' : mentees})

        return current_mentees
