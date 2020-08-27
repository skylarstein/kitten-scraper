from abc import ABCMeta, abstractmethod
from kitten_utils import Utils

class SheetReaderBase(metaclass=ABCMeta):
    def __init__(self):
        self._config_sheet_name = 'Config'
        self._mentor_sheets = []
        self._mentor_match_values = {}

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

    def _find_column_by_name(self, cells, name):
        for n in range(0, len(cells[0])):
            if cells[0][n].value == name:
                return n
        return 0

    def _is_reserved_sheet(self, sheet_name):
        return sheet_name.lower() in ['contact info',
                                      'config',
                                      'announcements',
                                      'resources',
                                      'calendar',
                                      'meetings/orientations dates',
                                      self._config_sheet_name.lower()]
