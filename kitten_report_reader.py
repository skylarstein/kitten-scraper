import re
import xlrd
from datetime import datetime
from kitten_utils import *

class KittenReportReader:
    ''' KittenReportReader will process process the incoming daily report, extend data with
        additional provided details, and output new results to csv
    '''

    def open_xls(self, xls_filename):
        ''' Open the daily report xls, perform some basic validation to make sure everything
            is in place.
        '''
        try:
            self.workbook = xlrd.open_workbook(xls_filename)
            self.sheet = self.workbook.sheet_by_index(0)

            if (self.sheet.row_values(0)[0] != 'Datetime of Current Status Date' or
                self.sheet.row_values(0)[1] != 'Current Animal Type' or
                self.sheet.row_values(0)[2] != 'AnimalID' or
                self.sheet.row_values(0)[3] != 'Animal Name' or
                self.sheet.row_values(0)[4] != 'Age' or
                self.sheet.row_values(0)[5] != 'Foster Parent ID'):

                print('ERROR: Unexpected column layout in {}'.format(xls_filename))
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
        for row_number in range(1, self.sheet.nrows):
            animal_number = self.sheet.row_values(row_number)[2]
            if isinstance(animal_number, float): # xls stores all numbers as float
                animal_numbers.add((int(animal_number)))

        return animal_numbers

    def xlsfloat_as_datetime(self, xlsfloat, workbook_datemode):
        ''' Convert Excel float date type to datetime
        '''
        if not xlsfloat:
            return None

        return datetime(*xlrd.xldate_as_tuple(xlsfloat, workbook_datemode))

    def copy_row_as_text(self, row_number):
        ''' Output is written as csv so we need to stringify all types (dates in particular)
        '''
        values = []
        for col_number in range(0, len(self.sheet.row_values(row_number))):
            cell_type = self.sheet.cell_type(row_number, col_number)

            if cell_type == xlrd.XL_CELL_DATE:
                dt = self.xlsfloat_as_datetime(self.sheet.row_values(row_number)[col_number], self.workbook.datemode)
                # wrapping datestr in ="%s" to deal with Excel auto-formatting issues
                values.append(dt.strftime('="%d-%b-%Y %-I:%M %p"'))

            elif cell_type == xlrd.XL_CELL_NUMBER:
                values.append(str(int(self.sheet.row_values(row_number)[col_number])))

            else:
                s = str(self.sheet.row_values(row_number)[col_number])
                if s == "null":
                    s = ''
                values.append('"{}"'.format(s))

        return values

    def pretty_print_animal_age(self, age_string):
        ''' Expecting an age string in the format '%d years %d months %d weeks'
        '''
        result = ''
        try:
            # For the sake of brevity in the spreadsheet, I'll shorten the age string when I can.
            # For example, if an animal is > 1 year old, there is no need to include months and weeks.
            #
            (years, months, weeks) = re.search(r'(\d+) years (\d+) months (\d+) weeks', age_string).groups()
            if int(years) > 0:
                result = '{} years'.format(years)
            elif int(months) >= 3:
                result = '{} months'.format(months)
            else:
                result = '{} weeks'.format(int(weeks) + int(months) * 4)
        except:
            result = 'Unknown Age'

        return result

    def count_animals(self, person_number, correct_foster_parents):
        ''' Count the number and age of each animal type assigned to this person number
        '''
        animal_types = []
        animal_numbers = {}
        animal_ages = {}
        last_animal_type = ''
        for row_number in range(1, self.sheet.nrows): # ignore header
            a_number = int(self.sheet.row_values(row_number)[2] or 0)
            if not a_number:
                continue

            a_type = self.sheet.row_values(row_number)[1]
            a_age = self.sheet.row_values(row_number)[4]
            p_number = next(p for p in correct_foster_parents if a_number in correct_foster_parents[p])

            if not a_type:
                a_type = last_animal_type
            else:
                a_type = a_type.encode('utf-8').strip()
                last_animal_type = a_type

            if person_number == p_number:
                animal_types.append(a_type)
                # WARNING: Making an assumption here that same animal types will be of the same age
                # (or at least close enough in grouped litters) so that choosing one age to share won't
                # be much of an issue.
                #
                if a_age:
                    animal_ages[a_type] = a_age

                if a_type in animal_numbers:
                    animal_numbers[a_type].append(a_number)
                else:
                    animal_numbers[a_type] = [a_number]

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
            age = self.pretty_print_animal_age(animal_ages[animal] if animal in animal_ages else '')
            numbers = ', '.join(str(a) for a in animal_numbers[animal])
            result_str += '{} {}{} @ {} ({})'.format(animal_counts[animal], animal, 's' if animal_counts[animal] > 1 else '', age, numbers)

        return result_str.lower()

    def output_results(self, persons_data, correct_foster_parents, csv_filename):
        ''' Combine the newly gathered person data with the daily report, output results
            to a new csv document.
        '''
        print_success('Writing results to {}...'.format(csv_filename))

        # First include the original column headers, then add columns for our new data
        #
        new_rows = []
        new_rows.append(self.sheet.row_values(0))

        new_rows[-1].append('Correct Foster Parent ID')
        new_rows[-1].append('Name')
        new_rows[-1].append('E-mail')
        new_rows[-1].append('Phone')
        new_rows[-1].append('Foster Experience')
        new_rows[-1].append('Date Kittens Received')	
        new_rows[-1].append('Quantity')
        new_rows[-1].append('Notes')

        for row_number in range(1, self.sheet.nrows): # ignore header
            animal_type = self.sheet.row_values(row_number)[1]
            animal_number = int(self.sheet.row_values(row_number)[2] or 0)
            #orig_person_number = int(self.sheet.row_values(row_number)[5])
            status_datetime = self.xlsfloat_as_datetime(self.sheet.row_values(row_number)[0], self.workbook.datemode)

            # If there is no animal number, skip the entire row
            #
            if not animal_number:
                continue

            # Include original column data as text since we're building a CSV document
            #
            new_rows.append(self.copy_row_as_text(row_number))

            corrected_person_number = next(p for p in correct_foster_parents if animal_number in correct_foster_parents[p])
            new_rows[-1].append('"{}"'.format(corrected_person_number))

            # Only include extended person details for rows with 'Current Animal Type' and 'Status Update'
            #
            if not animal_type or not status_datetime:
                continue

            # Grab the person data from the associated person number
            #
            person_data = persons_data[str(corrected_person_number)]
            name = person_data['full_name'] if 'preferred_name' in person_data else ''
            animal_quantity = self.count_animals(corrected_person_number, correct_foster_parents)
            prev_animals_fostered = person_data['prev_animals_fostered']
            notes = person_data['notes']
            foster_experience = 'NEW' if not prev_animals_fostered else '{} previous'.format(prev_animals_fostered)

            # Build phone number(s) string
            #
            cell_number = person_data['cell_phone'] if 'cell_phone' in person_data else ''
            home_number = person_data['home_phone'] if 'home_phone' in person_data else ''

            phone = ''
            if len(cell_number) >= 10: # ignore incomplete phone numbers
                phone = 'c: {}'.format(cell_number)

            if len(home_number) >= 10: # ignore incomplete phone numbers
                if len(phone):
                    phone += '\r'
                phone += 'h: {}'.format(home_number)

            # Build email(s) string
            #
            email = person_data['primary_email'] if 'primary_email' in person_data else ''
            secondary_email = person_data['secondary_email'] if 'secondary_email' in person_data else ''

            if len(secondary_email):
                if len(email):
                    email += '\r'
                email += secondary_email

            # Since we're receiving "last 24 hour reports" I'll assume received date is the same
            # day as the status date
            #
            date_received = status_datetime.strftime('%d-%b-%Y') if status_datetime else ''

            new_rows[-1].append('"{}"'.format(name))
            new_rows[-1].append('"{}"'.format(email))
            new_rows[-1].append('"{}"'.format(phone))
            new_rows[-1].append('"{}"'.format(foster_experience))
            new_rows[-1].append('="{}"'.format(date_received)) # using ="%s" for dates to deal with Excel auto-formatting issues
            new_rows[-1].append('"{}"'.format(animal_quantity))
            new_rows[-1].append('"{}"'.format(notes))

        with open(csv_filename, 'w') as outfile:
            for row in new_rows:
                outfile.write(','.join(row))
                outfile.write('\n')
