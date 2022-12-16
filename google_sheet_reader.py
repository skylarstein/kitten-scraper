from datetime import date
import time
import pygsheets
from kitten_utils import Log, Utils
from sheet_reader_base import SheetReaderBase

class GoogleSheetReader(SheetReaderBase):
    def load_mentors_spreadsheet(self, auth):
        ''' Load the feline foster spreadsheet
        '''
        start_time = time.time()
        try:
            Log.success(f'Loading mentors spreadsheet from Google Sheets (id = \'{auth["google_spreadsheet_key"]}\')...')

            client = pygsheets.authorize(auth['google_client_secret'])
            spreadsheet = client.open_by_key(auth['google_spreadsheet_key'])

            config_yaml = spreadsheet.worksheet_by_title(self._CONFIG_SHEET_NAME)[2][0]

            for worksheet in spreadsheet.worksheets():
                if not self._is_reserved_sheet(worksheet.title) and not worksheet.hidden:
                    Log.debug(f'Reading worksheet \"{worksheet.title}\"...')
                    try:
                        if self.check_for_surgery_sheet(worksheet):
                            continue

                        # Mentor sheet header rows vary slightly between feline and canine. Perform a terrible quick-and-dirty validation.
                        #
                        if ['ID'] not in worksheet.get_values('E1', 'E2'):
                            raise Exception('') from Exception

                        # Build a list of mentee names/emails/ids to be used for mentor matching
                        #
                        b_rows = worksheet.get_values('B2', f'B{worksheet.rows}', include_tailing_empty = False, include_tailing_empty_rows = False)
                        c_rows = worksheet.get_values('C2', f'C{worksheet.rows}', include_tailing_empty = False, include_tailing_empty_rows = False)
                        e_rows = worksheet.get_values('E2', f'E{worksheet.rows}', include_tailing_empty = False, include_tailing_empty_rows = False)

                        mentor_match_cells = b_rows + c_rows + e_rows
                        self._mentor_match_values[Utils.utf8(worksheet.title)] = [Utils.utf8(item).lower() for sublist in mentor_match_cells for item in sublist]
                        self._mentor_sheets.append(worksheet)

                    except Exception:
                        Log.debug(f'Sheet \'{worksheet.title}\' does not appear to be a mentor sheet (skipping)')

        except Exception as e:
            Log.error(f'ERROR: Unable to load mentors spreadsheet!\r\n{str(e)}, {repr(e)}')
            return None

        print('Loaded {0} mentors from \"{1}\" in {2:.0f} seconds'.format(len(self._mentor_sheets), spreadsheet.title, time.time() - start_time))
        return config_yaml

    def check_for_surgery_sheet(self, worksheet):
        if any(worksheet.title in substr for substr in self._SURGERY_SHEET_NAMES):
            surgery_rows = worksheet.get_values('A1', f'H{worksheet.rows}', include_tailing_empty = False, include_tailing_empty_rows = False)
            date_col = -1
            patient_col = -1
            for col in range(0, len(surgery_rows[0])):
                # Allow for an extra header row (accounting for differences between Feline and Canine)
                #
                if any('date' in substr for substr in [str(surgery_rows[0][col]).lower(), str(surgery_rows[1][col]).lower()]):
                    date_col = col
                elif 'patient' in (str(surgery_rows[0][col]).lower(), str(surgery_rows[1][col]).lower()):
                    patient_col = col

            if date_col != -1 and patient_col != -1:
                for row in range(1, len(surgery_rows)):
                    try:
                        a_number = surgery_rows[row][patient_col]
                        if a_number.isdigit():
                            a_number = int(a_number)
                            # If there are multiple entries for a given a_number, assume the first is the most recent.
                            #
                            if a_number not in self._surgery_dates:
                                self._surgery_dates[int(a_number)] = surgery_rows[row][date_col]
                    except Exception as e:
                        Log.warn(f'{worksheet.title} column {patient_col}, row {row} is empty. Assuming this is the end of the list.')
                        break
            else:
                Log.error(f'Surgery form is not in expected format (date_col={date_col}, patient_col={patient_col}. Skipping.')

            Log.debug(f'Loaded {len(self._surgery_dates)} entries from the surgery sheet')
            return True

        return False

    def get_current_mentees(self):
        ''' Return the current mentees assigned to each mentor
        '''
        current_mentees = []
        for worksheet in self._mentor_sheets:
            if worksheet.title.lower() == 'retired mentor':
                continue
            print(f'Loading current mentees for {worksheet.title}... ', end='', flush=True)

            # It's much faster to grab a whole block of cells at once vs iterating through many API calls
            #
            max_search_rows = min(100, worksheet.rows)
            cells = worksheet.range(f'A1:G{max_search_rows}', returnas='cells')

            name_col_id = self._find_column_by_name(cells, 'Name')
            pid_col_id = self._find_column_by_name(cells, 'ID')
            date_col_id = self._find_column_by_name(cells, 'Date\nKittens\nReceived')
            if date_col_id == -1:
                date_col_id = self._find_column_by_name(cells, 'Date Dog Received')

            mentees = []
            search_failed = False
            most_recent_received_date = None
            for i in range(1, max_search_rows):
                if i == max_search_rows - 1:
                    search_failed = True
                    Log.error(f'Unable to determine current mentees for mentor {worksheet.title}')
                    mentees = []
                    break

                elif str(cells[i][0].value).lower().find('completed mentees') >= 0:
                    break # We've reached the end of the "active mentee" rows

                elif cells[i][name_col_id].value and str(cells[i][pid_col_id].value).isdigit():
                    mentee_name = cells[i][name_col_id].value
                    pid = int(cells[i][pid_col_id].value)
                    received_date = Utils.string_to_datetime(cells[i][date_col_id].value)

                    if received_date and (most_recent_received_date is None or received_date > most_recent_received_date):
                        most_recent_received_date = received_date

                    if not [mentee for mentee in mentees if mentee['pid'] == pid]: # ignore duplicate mentees
                        mentees.append({'name' : mentee_name, 'pid' : pid})

            if not search_failed:
                print(f'found {len(mentees)}')

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
                cells = worksheet.range(f'A1:G{max_search_rows}', returnas='cells')

                name_col_id = self._find_column_by_name(cells, 'Name')
                pid_col_id = self._find_column_by_name(cells, 'ID')
                notes_col_id = 0

                for i in range(1, max_search_rows):
                    if str(cells[i][0].value).lower().find('completed mentees') >= 0:
                        break # We've reached the end of the "active mentee" rows

                    if cells[i][name_col_id].value and str(cells[i][pid_col_id].value).isdigit():
                        pid = int(cells[i][pid_col_id].value)
                        if pid in mentee_ids:
                            # If this mentee name cell is already marked with strikethrough, leave it alone
                            #
                            name_cell_format = cells[i][name_col_id].text_format
                            if not name_cell_format or 'strikethrough' not in name_cell_format or name_cell_format['strikethrough'] is False:
                                mentee_name = cells[i][name_col_id].value
                                mentee_name = mentee_name.replace('\n', ' ').replace('\r', '')
                                Log.debug(f'Completed: {mentee_name} ({pid}) @ {mentor}[\'{cells[i][name_col_id].label}\']')
                                debug_mode = False
                                if not debug_mode:
                                    cells[i][name_col_id].set_text_format('strikethrough', True)
                                    notes_current_value = cells[i][notes_col_id].value
                                    if 'autoupdate: no animals' not in notes_current_value.lower():
                                        cells[i][notes_col_id].set_value(f'AutoUpdate: No animals {date.today().strftime("%b %-d, %Y")}\r\n{notes_current_value}')
