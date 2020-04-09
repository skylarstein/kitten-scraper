import xlrd
from datetime import datetime
from kitten_utils import *

class KittenReportReader(object):
    ''' KittenReportReader will load animal numbers from the incoming daily foster report. All other data will be
        queried during runtime since other data in the report may already be outdated (or even initially incorrect).
    '''
    def read_animal_numbers_from_xls(self, xls_filename):
        ''' Open the daily report xls, read animal numbers
        '''
        try:
            self._workbook = xlrd.open_workbook(xls_filename)
            self._sheet = self._workbook.sheet_by_index(0)

            if not self._sheet.nrows:
                print_err('ERROR: I\'m afraid you have an empty report: {}'.format(xls_filename))
                return None

            # Perform some initial sanity checks
            #
            STATUS_DATE_COL = 0
            ANIMAL_TYPE_COL = 1
            ANIMAL_ID_COL = 2
            ANIMAL_NAME_COL = 3
            ANIMAL_AGE_COL = 4
            FOSTER_PARENT_ID_COL = 5

            if (self._sheet.row_values(0)[STATUS_DATE_COL] != 'Datetime of Current Status Date' or
                self._sheet.row_values(0)[ANIMAL_TYPE_COL] != 'Current Animal Type' or
                self._sheet.row_values(0)[ANIMAL_ID_COL] != 'AnimalID' or
                self._sheet.row_values(0)[ANIMAL_NAME_COL] != 'Animal Name' or
                self._sheet.row_values(0)[ANIMAL_AGE_COL] != 'Age' or
                self._sheet.row_values(0)[FOSTER_PARENT_ID_COL] != 'Foster Parent ID'):

                print_err('ERROR: Unexpected column layout in the report. Something has changed! {}'.format(xls_filename))
                return None

            print_success('Loaded report {}'.format(xls_filename))

            animal_numbers = set()
            for row_number in range(1, self._sheet.nrows):
                animal_number = self._sheet.row_values(row_number)[ANIMAL_ID_COL]
                # xls stores all numbers as float, but also handle str type just in case
                #
                if isinstance(animal_number, float) or (isinstance(animal_number, str) and animal_number.isdigit()):
                    animal_numbers.add((int(animal_number)))

            return animal_numbers

        except IOError as err:
            print_err('ERROR: Unable to read xls file: {}, {}'.format(xls_filename, err))

        except xlrd.XLRDError as err:
            print_err('ERROR: Unable to read xls file: {}, {}'.format(xls_filename, err))

        return None

    def _xlsfloat_as_datetime(self, xlsfloat, workbook_datemode):
        if not xlsfloat:
            return None

        return datetime(*xlrd.xldate_as_tuple(xlsfloat, workbook_datemode))

    def _cell_to_string(self, row_number, col_number):
        cell_type = self._sheet.cell_type(row_number, col_number)
        result = ''

        if cell_type == xlrd.XL_CELL_DATE:
            dt = self._xlsfloat_as_datetime(self._sheet.row_values(row_number)[col_number], self._workbook.datemode)
            result = dt.strftime('="%d-%b-%Y %-I:%M %p"')

        elif cell_type == xlrd.XL_CELL_NUMBER:
            result = '="{}"'.format(str(int(self._sheet.row_values(row_number)[col_number])))

        else:
            s = str(self._sheet.row_values(row_number)[col_number])
            result = '"{}"'.format(s if s != 'null' else '')

        return result
