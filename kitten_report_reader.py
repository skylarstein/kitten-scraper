import re
import xlrd
from datetime import datetime
from kitten_utils import *

class KittenReportReader(object):
    ''' KittenReportReader will process process the incoming daily report, extend data with
        additional provided details, and output new results to csv
    '''

    def __init__(self):
        self.__STATUS_DATE_COL = 0
        self.__ANIMAL_TYPE_COL = 1
        self.__ANIMAL_ID_COL = 2
        self.__ANIMAL_NAME_COL = 3
        self.__ANIMAL_AGE_COL = 4
        self.__FOSTER_PARENT_ID_COL = 5

    def open_xls(self, xls_filename):
        ''' Open the daily report xls, perform some basic validation to make sure everything is in place.
        '''
        try:
            self._workbook = xlrd.open_workbook(xls_filename)
            self._sheet = self._workbook.sheet_by_index(0)

            if not self._sheet.nrows:
                print_err('ERROR: I\'m afraid you have an empty report: {}'.format(xls_filename))
                return False

            if (self._sheet.row_values(0)[self.__STATUS_DATE_COL] != 'Datetime of Current Status Date' or
                self._sheet.row_values(0)[self.__ANIMAL_TYPE_COL] != 'Current Animal Type' or
                self._sheet.row_values(0)[self.__ANIMAL_ID_COL] != 'AnimalID' or
                self._sheet.row_values(0)[self.__ANIMAL_NAME_COL] != 'Animal Name' or
                self._sheet.row_values(0)[self.__ANIMAL_AGE_COL] != 'Age' or
                self._sheet.row_values(0)[self.__FOSTER_PARENT_ID_COL] != 'Foster Parent ID'):

                print_err('ERROR: Unexpected column layout in the report. Something unexpected has changed! {}'.format(xls_filename))
                return False

            print_success('Loaded report {}'.format(xls_filename))
            return True

        except IOError as err:
            print_err('ERROR: Unable to read xls file: {}, {}'.format(xls_filename, err))

        except xlrd.XLRDError as err:
            print_err('ERROR: Unable to read xls file: {}'.format(err.message))

        return False

    def get_animal_numbers(self):
        ''' Return a set of animal numbers found in the daily report
        '''
        animal_numbers = set()
        for row_number in range(1, self._sheet.nrows):
            animal_number = self._sheet.row_values(row_number)[self.__ANIMAL_ID_COL]
            if isinstance(animal_number, float): # xls stores all numbers as float
                animal_numbers.add((int(animal_number)))

        return animal_numbers

    def output_results(self, persons_data, correct_foster_parents, animal_special_messages, sn_status, csv_filename):
        ''' Combine the newly gathered person data with the daily report, output results
            to a new csv document.
        '''
        print_success('Writing results to {}...'.format(csv_filename))

        # First include the original column headers, then add columns for our new data
        #
        new_rows = []
        new_rows.append(self._sheet.row_values(0))

        new_rows[-1].append('Correct Foster Parent ID')
        new_rows[-1].append('Report Notes')
        new_rows[-1].append('Name')
        new_rows[-1].append('E-mail')
        new_rows[-1].append('Phone')
        new_rows[-1].append('Foster Experience')
        new_rows[-1].append('Date Kittens Received')
        new_rows[-1].append('Quantity')
        new_rows[-1].append('Special Animal Message')

        processed_p_numbers = []
        for row_number in range(1, self._sheet.nrows): # ignore header
            animal_type = self._sheet.row_values(row_number)[self.__ANIMAL_TYPE_COL]
            animal_number = int(self._sheet.row_values(row_number)[self.__ANIMAL_ID_COL] or 0)
            status_datetime = self._xlsfloat_as_datetime(self._sheet.row_values(row_number)[self.__STATUS_DATE_COL], self._workbook.datemode)

            # If there is no animal number, skip the entire row
            #
            if not animal_number:
                continue

            # Include original column data as text since we're building a CSV document
            #
            new_rows.append(self._copy_row_as_text(row_number))

            # Only include an extended details row once per foster parent
            #
            corrected_person_number = next((p for p in correct_foster_parents if animal_number in correct_foster_parents[p]), 'UNKNOWN')

            if corrected_person_number in processed_p_numbers:
                continue
            else:
                processed_p_numbers.append(corrected_person_number)

            # Grab the person data from the associated person number
            #
            person_data = persons_data[corrected_person_number] if corrected_person_number in persons_data else {}
            name = person_data['full_name'] if 'full_name' in person_data else ''

            animal_quantity_string, animal_numbers = self._count_animals(corrected_person_number, correct_foster_parents, sn_status)
            prev_animals_fostered = person_data['prev_animals_fostered'] if 'prev_animals_fostered' in person_data else None
            report_notes = person_data['notes'] if 'notes' in person_data else ''

            special_message = ''
            for a in animal_numbers:
                msg = animal_special_messages[a]
                if msg:
                    special_message += '{}{}: {}'.format('\r\r' if special_message else '', a, animal_special_messages[a])

            if prev_animals_fostered is not None:
                foster_experience = 'NEW' if not prev_animals_fostered else '{} previous'.format(prev_animals_fostered)
            else:
                foster_experience = ''

            # Build phone number(s) string
            #
            cell_number = person_data['cell_phone'] if 'cell_phone' in person_data else ''
            home_number = person_data['home_phone'] if 'home_phone' in person_data else ''

            phone = ''
            if len(cell_number) >= 10: # ignore incomplete phone numbers
                phone = 'c: {}'.format(cell_number)

            if len(home_number) >= 10: # ignore incomplete phone numbers
                phone += '{}h: {}'.format('\r' if phone else '', home_number)

            # Build email(s) string
            #
            email = person_data['primary_email'] if 'primary_email' in person_data else ''
            secondary_email = person_data['secondary_email'] if 'secondary_email' in person_data else ''

            if secondary_email:
                email += '{}{}'.format('\r' if email else '', secondary_email)

            # Since the reports cover "last 24 hours" I'll assume received date is the same day as the status date
            #
            date_received = status_datetime.strftime('%d-%b-%Y') if status_datetime else ''

            new_rows[-1].append('"{}"'.format(corrected_person_number))
            new_rows[-1].append('"{}"'.format(report_notes))
            new_rows[-1].append('"{}"'.format(name))
            new_rows[-1].append('"{}"'.format(email))
            new_rows[-1].append('"{}"'.format(phone))
            new_rows[-1].append('"{}"'.format(foster_experience))
            new_rows[-1].append('="{}"'.format(date_received)) # using ="%s" for dates to deal with Excel auto-formatting issues
            new_rows[-1].append('"{}"'.format(animal_quantity_string))
            new_rows[-1].append('"{}"'.format(special_message))

            print('{} = {}'.format(name, animal_numbers))

        with open(csv_filename, 'w') as outfile:
            for row in new_rows:
                outfile.write(','.join(row))
                outfile.write('\n')

    def _xlsfloat_as_datetime(self, xlsfloat, workbook_datemode):
        ''' Convert Excel float date type to datetime
        '''
        if not xlsfloat:
            return None

        return datetime(*xlrd.xldate_as_tuple(xlsfloat, workbook_datemode))

    def _copy_row_as_text(self, row_number):
        ''' Output is written as csv so we need to stringify all types (dates in particular)
        '''
        values = []
        for col_number in range(0, len(self._sheet.row_values(row_number))):
            cell_type = self._sheet.cell_type(row_number, col_number)

            if cell_type == xlrd.XL_CELL_DATE:
                dt = self._xlsfloat_as_datetime(self._sheet.row_values(row_number)[col_number], self._workbook.datemode)
                # wrapping datestr in ="%s" to deal with Excel auto-formatting issues
                values.append(dt.strftime('="%d-%b-%Y %-I:%M %p"'))

            elif cell_type == xlrd.XL_CELL_NUMBER:
                values.append(str(int(self._sheet.row_values(row_number)[col_number])))

            else:
                s = str(self._sheet.row_values(row_number)[col_number])
                values.append('"{}"'.format(s if s != 'null' else ''))

        return values

    def _pretty_print_animal_age(self, age_string):
        ''' Expecting an age string in the format '%d years %d months %d weeks'
        '''
        pretty_age_string = ''
        animal_type = ''
        try:
            # For the sake of brevity in the spreadsheet, I'll shorten the age string when I can.
            # For example, if an animal is > 1 year old, there's no huge need to include months and weeks.
            # I'll also return an animal_type string ['kitten', 'cat'] based on age since daily reports
            # occasionally fail to include this information.
            #
            (years, months, weeks) = re.search(r'(\d+) years (\d+) months (\d+) weeks', age_string).groups()
            if int(years) > 0:
                pretty_age_string = '{} years'.format(years)
                animal_type = 'cat'
            elif int(months) >= 3:
                pretty_age_string = '{} months'.format(months)
                animal_type = 'kitten' if int(months) <= 6 else 'cat'
            else:
                pretty_age_string = '{} weeks'.format(int(weeks) + int(months) * 4)
                animal_type = 'kitten'
        except:
            animal_type = 'kitten' # I have no idea, but 'kitten' is a safe bet
            pretty_age_string = 'Unknown Age'

        return pretty_age_string, animal_type

    def _count_animals(self, person_number, correct_foster_parents, sn_status):
        ''' Count the number and age of each animal type assigned to this person number
        '''
        animal_types = []
        animals_by_type = {}
        animal_numbers = []
        animal_ages = {}
        last_animal_type = ''
        for row_number in range(1, self._sheet.nrows): # ignore header
            a_number = int(self._sheet.row_values(row_number)[self.__ANIMAL_ID_COL] or 0)
            if not a_number:
                continue

            a_type = self._sheet.row_values(row_number)[self.__ANIMAL_TYPE_COL]
            a_age = self._sheet.row_values(row_number)[self.__ANIMAL_AGE_COL]
            p_number = next((p for p in correct_foster_parents if a_number in correct_foster_parents[p]), None)

            if not a_type:
                a_type = last_animal_type
            else:
                a_type = utf8(a_type)
                last_animal_type = a_type

            if person_number == p_number:
                animal_types.append(a_type)
                animal_numbers.append(a_number)

                # WARNING: Making an assumption here that same animal types will be of the same age
                # (or at least close enough in grouped litters) so that choosing one age to share won't
                # be much of an issue.
                #
                if a_age:
                    animal_ages[a_type] = a_age

                animals_by_type.setdefault(a_type, []).append(a_number)

        # We now have a list of all animal types. Next, create a set with total counts per type.
        #
        animal_counts = {}
        for animal in set(animal_types):
            animal_counts[animal] = animal_types.count(animal)

        # Pretty-print results
        #
        result_str = ''
        for animal in animal_counts:
            if result_str:
                result_str += '\r'
            age, animal_type = self._pretty_print_animal_age(animal_ages[animal] if animal in animal_ages else '')

            # Include spay/neuter status with each animal number
            a_numbers_str = ''
            for a in animals_by_type[animal]:
                a_numbers_str += '{}{} (SN={})'.format('\r' if a_numbers_str else '', a, sn_status[a] if a in sn_status else 'Unknown')

            result_str += '{} {}{} @ {}\r{}'.format(animal_counts[animal], animal_type.lower(), 's' if animal_counts[animal] > 1 else '', age, a_numbers_str)

        return result_str, animal_numbers
