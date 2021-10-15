import os
import xlrd
from boxsdk import Client, JWTAuth
from kitten_utils import Log, Utils
from sheet_reader_base import SheetReaderBase

class BoxSheetReader(SheetReaderBase):
    def load_mentors_spreadsheet(self, auth):
        ''' Load the feline foster spreadsheet
        '''
        try:
            Log.success(f'Loading mentors spreadsheet from Box (id = {auth["box_file_id"]})...')

            jwt_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), auth['box_jwt'])
            client = Client(JWTAuth.from_settings_file(jwt_path))
            box_file = client.as_user(client.user(user_id = auth['box_user_id'])).file(file_id=auth['box_file_id']).get()
            xlxs_workbook = xlrd.open_workbook(file_contents=box_file.content())

            config_yaml = xlxs_workbook.sheet_by_name(self._CONFIG_SHEET_NAME).row_values(1)[0]

            for sheet_name in xlxs_workbook.sheet_names():
                if not self._is_reserved_sheet(sheet_name):
                    sheet = xlxs_workbook.sheet_by_name(sheet_name)
                    self._mentor_sheets.append(sheet)
                    all_values = [sheet.row_values(i) for i in range(1, sheet.nrows)]
                    self._mentor_match_values[Utils.utf8(sheet_name)] = [Utils.utf8(str(item)).lower() for sublist in all_values for item in sublist]

        except Exception as e:
            Log.error(f'ERROR: Unable to load Feline Foster spreadsheet!\r\n{str(e)}, {repr(e)}')
            return None

        print(f'Loaded {len(self._mentor_sheets)} mentors from \"{box_file["name"]}\"')
        return config_yaml

    def get_current_mentees(self):
        ''' Return the current mentees assigned to each mentor
        '''
        current_mentees = []
        for worksheet in self._mentor_sheets:
            if worksheet.name.lower() == 'retired mentor':
                continue

            print(f'Loading current mentees for {worksheet.name}... ', end='')

            # It's much faster to grab a whole block of cells at once vs iterating through many API calls
            #
            max_search_rows = min(50, worksheet.nrows)
            cells = [worksheet.row_slice(row, start_colx=0, end_colx=7) for row in range(0, max_search_rows)]

            name_col_id = self._find_column_by_name(cells, 'Name')
            pid_col_id = self._find_column_by_name(cells, 'ID')

            mentees = []
            search_failed = False
            for i in range(1, max_search_rows):
                if i == max_search_rows - 1:
                    search_failed = True
                    Log.error(f'Unable to determine current mentees for mentor {worksheet.name}')
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
                print(f'found {len(mentees)}')

            current_mentees.append({ 'mentor' : worksheet.name, 'mentees' : mentees})

        return current_mentees

    def set_completed_mentees(self, mentor, mentee_ids):
        raise NotImplementedError(f'set_completed_mentees not implemented in {__file__}')
