from datetime import date
import pygsheets
from kitten_utils import Log, Utils
from sheet_reader_base import SheetReaderBase

class GoogleSheetReader(SheetReaderBase):
    def load_mentors_spreadsheet(self, auth):
        ''' Load the feline foster spreadsheet
        '''
        try:
            Log.success('Loading mentors spreadsheet from Google Sheets (id = \'{}\')...'.format(auth['google_spreadsheet_key']))

            client = pygsheets.authorize(auth['google_client_secret'])
            spreadsheet = client.open_by_key(auth['google_spreadsheet_key'])

            config_yaml = spreadsheet.worksheet_by_title(self._config_sheet_name)[2][0]

            for worksheet in spreadsheet.worksheets():
                if not self._is_reserved_sheet(worksheet.title):
                    self._mentor_sheets.append(worksheet)

                    try:
                        # Build a list of mentee names/emails/ids to be used for mentor matching
                        #
                        mentor_match_cells = worksheet.get_values('B2', 'B{}'.format(worksheet.rows), include_tailing_empty = False, include_tailing_empty_rows = False)
                        mentor_match_cells += worksheet.get_values('C2', 'C{}'.format(worksheet.rows), include_tailing_empty = False, include_tailing_empty_rows = False)
                        mentor_match_cells += worksheet.get_values('E2', 'E{}'.format(worksheet.rows), include_tailing_empty = False, include_tailing_empty_rows = False)
                        self._mentor_match_values[Utils.utf8(worksheet.title)] = [Utils.utf8(item).lower() for sublist in mentor_match_cells for item in sublist]

                        # Header rows vary slightly between feline and canine. Perform a quick and dirty validation.
                        #
                        if ['ID'] not in worksheet.get_values('E1', 'E2'):
                            raise Exception('') from Exception

                    except Exception:
                        Log.debug('Sheet \'{}\' does not appear to be a mentor sheet (skipping)'.format(worksheet.title))

        except Exception as e:
            Log.error('ERROR: Unable to load mentors spreadsheet!\r\n{}, {}'.format(str(e), repr(e)))
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
            print('Loading current mentees for {}... '.format(worksheet.title), end='', flush=True)

            # It's much faster to grab a whole block of cells at once vs iterating through many API calls
            #
            max_search_rows = min(100, worksheet.rows)
            cells = worksheet.range('A1:G{}'.format(max_search_rows), returnas='cells')

            name_col_id = self._find_column_by_name(cells, 'Name')
            pid_col_id = self._find_column_by_name(cells, 'ID')
            date_col_id = self._find_column_by_name(cells, 'Date\nKittens\nReceived')

            mentees = []
            search_failed = False
            most_recent_received_date = None
            for i in range(1, max_search_rows):
                if i == max_search_rows - 1:
                    search_failed = True
                    Log.error('Unable to determine current mentees for mentor {}'.format(worksheet.title))
                    mentees = []
                    break

                elif str(cells[i][0].value).lower().find('completed mentees') >= 0:
                    break # We've reached the end of the "active mentee" rows

                elif cells[i][name_col_id].value and cells[i][pid_col_id].value:
                    mentee_name = cells[i][name_col_id].value
                    pid = int(cells[i][pid_col_id].value)
                    received_date = Utils.string_to_datetime(cells[i][date_col_id].value)

                    if received_date and (most_recent_received_date is None or received_date > most_recent_received_date):
                        most_recent_received_date = received_date

                    if not [mentee for mentee in mentees if mentee['pid'] == pid]: # ignore duplicate mentees
                        mentees.append({'name' : mentee_name, 'pid' : pid})

            if not search_failed:
                print('found {}'.format(len(mentees)))

            current_mentees.append({ 'mentor' : worksheet.title,
                                     'mentees' : mentees,
                                     'most_recent' : most_recent_received_date})
        return current_mentees

    def set_completed_mentees(self, mentor, mentee_ids):
        ''' Mark the given mentees as completed.

            Future refactoring consideration: See similar code between set_completed_mentees() and get_current_mentees().
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
                            # If this mentee name cell is already marked with strikethrough, leave it alone
                            #
                            name_cell_format = cells[i][name_col_id].text_format
                            if not name_cell_format or 'strikethrough' not in name_cell_format or name_cell_format['strikethrough'] is False:
                                mentee_name = cells[i][name_col_id].value
                                mentee_name = mentee_name.replace('\n', ' ').replace('\r', '')
                                Log.debug('Completed: {} ({}) @ {}[\'{}\']'.format(mentee_name, pid, mentor, cells[i][name_col_id].label))
                                cells[i][name_col_id].set_text_format('strikethrough', True)
                                current_value = cells[i][0].value
                                if 'autoupdate: no animals' not in current_value.lower():
                                    cells[i][0].set_value('AutoUpdate: No animals {}\r\n{}'.format(date.today().strftime('%b %-d, %Y'), current_value))
