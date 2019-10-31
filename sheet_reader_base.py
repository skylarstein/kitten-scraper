from abc import ABCMeta, abstractmethod

class SheetReaderBase(metaclass=ABCMeta):
    def __init__(self):
        self._config_sheet_name = 'Config'

    @abstractmethod
    def load_mentors_spreadsheet(self, auth):
        pass

    @abstractmethod
    def find_matching_mentors(self, match_strings):
        pass

    @abstractmethod
    def get_current_mentees(self):
        pass

    def _is_reserved_sheet(self, sheet_name):
        return sheet_name.lower() in ['contact info',
                                      'updates',
                                      'announcements',
                                      'resources',
                                      'calendar',
                                      self._config_sheet_name.lower()]
