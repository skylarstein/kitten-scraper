from abc import ABCMeta, abstractmethod

class SheetReaderBase(metaclass=ABCMeta):
    @abstractmethod
    def load_mentors_spreadsheet(self, auth):
        pass

    @abstractmethod
    def find_matching_mentors(self, match_strings):
        pass

    @abstractmethod
    def get_current_mentees(self):
        pass
