from abc import ABCMeta, abstractmethod
from kitten_utils import Utils

class SheetReaderBase(metaclass=ABCMeta):
    def __init__(self):
        self._CONFIG_SHEET_NAME = 'Config'
        self._SURGERY_SHEET_NAMES = ['Foster S-N Appts', 'Spay_Neuter Appts']
        self._mentor_sheets = []
        self._mentor_match_values = {}
        self._surgery_dates = {}

    @abstractmethod
    def load_mentors_spreadsheet(self, auth):
        pass

    @abstractmethod
    def get_current_mentees(self):
        pass

    @abstractmethod
    def set_completed_mentees(self, mentor, mentee_ids):
        pass

    def find_matching_mentors(self, match_strings):
        ''' Find mentor worksheets that match any string in match_strings. Not very sophisticated right now,
            I'm simply searching for a match anywhere in each mentor sheet.
        '''
        match_strings = [Utils.utf8(s).lower() for s in match_strings if s]
        matching_mentors = set()

        for sheet_name in self._mentor_match_values:
            if [item for item in self._mentor_match_values[sheet_name] if any(match in item for match in match_strings)]:
                matching_mentors.add(sheet_name)

        return matching_mentors

    def get_surgery_date(self, a_number):
        return self._surgery_dates[a_number] if a_number in self._surgery_dates else ''

    def _find_column_by_name(self, cells, name):
        for n in range(0, len(cells[0])):
            # Allow for an optional header when searching for column names
            if any(name in substr for substr in [cells[0][n].value, cells[1][n].value]):
                return n
        return -1

    def _is_reserved_sheet(self, sheet_name):
        return sheet_name.lower() in ['contact info',
                                      'config',
                                      'announcements',
                                      'resources',
                                      'calendar',
                                      'template',
                                      'forms',
                                      'mentor meetings',
                                      'meetings/orientations dates',
                                      self._CONFIG_SHEET_NAME.lower()]
