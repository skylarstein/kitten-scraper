from datetime import date
import os
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

            config_yaml = spreadsheet.worksheet_by_title(self._config_sheet_name)[2][0]

            for worksheet in spreadsheet.worksheets():
                if not self._is_reserved_sheet(worksheet.title):
                    self._mentor_sheets.append(worksheet)

                    # Build a list of mentee names/emails/ids to be used for mentor matching
                    #
                    mentor_match_cells = worksheet.get_values('B2', 'B{}'.format(worksheet.rows), include_tailing_empty = False, include_tailing_empty_rows = False)
                    mentor_match_cells += worksheet.get_values('C2', 'C{}'.format(worksheet.rows), include_tailing_empty = False, include_tailing_empty_rows = False)
                    mentor_match_cells += worksheet.get_values('E2', 'E{}'.format(worksheet.rows), include_tailing_empty = False, include_tailing_empty_rows = False)
                    self._mentor_match_values[utf8(worksheet.title)] = [utf8(item).lower() for sublist in mentor_match_cells for item in sublist]

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
            max_search_rows = min(100, worksheet.rows)
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
                    break # We've reached the end of the "active mentee" rows

                elif cells[i][name_col_id].value and cells[i][pid_col_id].value:
                    mentee_name = cells[i][name_col_id].value
                    pid = int(cells[i][pid_col_id].value)
                    if not [mentee for mentee in mentees if mentee['pid'] == pid]: # ignore duplicate mentees
                        mentees.append({'name' : mentee_name, 'pid' : pid})

            if not search_failed:
                print('found {}'.format(len(mentees)))

            current_mentees.append({ 'mentor' : worksheet.title, 'mentees' : mentees})

        return current_mentees

    def set_completed_mentees(self, mentor, mentee_ids):
        ''' Mark the given mentees as completed.

            Future refactoring consideration: Some similar code between set_completed_mentees() and get_current_mentees().
        '''
        for worksheet in self._mentor_sheets:
            if worksheet.title.lower() == mentor.lower():
                max_search_rows = min(100, worksheet.rows)
                cells = worksheet.range('A1:G{}'.format(max_search_rows), returnas='cells')

                name_col_id = self._find_column_by_name(cells, 'Name')
                pid_col_id = self._find_column_by_name(cells, 'ID')

                for i in range(1, max_search_rows):
                    if str(cells[i][0].value).lower().find('completed mentees') >= 0:
                        break # We've reached the end of the "active mentee" rows

                    elif cells[i][name_col_id].value and cells[i][pid_col_id].value:
                        pid = int(cells[i][pid_col_id].value)
                        if pid in mentee_ids:
                            mentee_name = cells[i][name_col_id].value
                            print_debug("Completed: {} ({}) @ {}[{}]".format(mentee_name, pid, mentor, i))
                            cells[i][name_col_id].set_text_format('strikethrough', True)
                            current_value = cells[i][0].value
                            if 'autoupdate: no animals' not in current_value.lower():
                                cells[i][0].set_value('AutoUpdate: No animals {}\r\n{}'.format(date.today().strftime('%Y.%m.%d'), current_value))
