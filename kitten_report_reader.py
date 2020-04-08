import re
import xlrd
from datetime import datetime
from kitten_utils import *

class KittenReportReader(object):
    ''' KittenReportReader will process the incoming daily report, query additional details, then output new results
    '''
    def __init__(self, dog_mode):
        self.STATUS_DATE_COL = 0
        self.ANIMAL_TYPE_COL = 1
        self.ANIMAL_ID_COL = 2
        self.ANIMAL_NAME_COL = 3
        self.ANIMAL_AGE_COL = 4
        self.FOSTER_PARENT_ID_COL = 5

        self.dog_mode = dog_mode
        self.ADULT_ANIMAL_TYPE = 'cat' if not dog_mode else 'dog'
        self.YOUNG_ANIMAL_TYPE = 'kitten' if not dog_mode else 'puppy'

    def open_xls(self, xls_filename):
        ''' Open the daily report xls, perform some basic sanity checks
        '''
        try:
            self._workbook = xlrd.open_workbook(xls_filename)
            self._sheet = self._workbook.sheet_by_index(0)

            if not self._sheet.nrows:
                print_err('ERROR: I\'m afraid you have an empty report: {}'.format(xls_filename))
                return False

            if (self._sheet.row_values(0)[self.STATUS_DATE_COL] != 'Datetime of Current Status Date' or
                self._sheet.row_values(0)[self.ANIMAL_TYPE_COL] != 'Current Animal Type' or
                self._sheet.row_values(0)[self.ANIMAL_ID_COL] != 'AnimalID' or
                self._sheet.row_values(0)[self.ANIMAL_NAME_COL] != 'Animal Name' or
                self._sheet.row_values(0)[self.ANIMAL_AGE_COL] != 'Age' or
                self._sheet.row_values(0)[self.FOSTER_PARENT_ID_COL] != 'Foster Parent ID'):

                print_err('ERROR: Unexpected column layout in the report. Something has changed! {}'.format(xls_filename))
                return False

            print_success('Loaded report {}'.format(xls_filename))
            return True

        except IOError as err:
            print_err('ERROR: Unable to read xls file: {}, {}'.format(xls_filename, err))

        except xlrd.XLRDError as err:
            print_err('ERROR: Unable to read xls file: {}, {}'.format(xls_filename, err))

        return False

    def get_animal_numbers(self):
        ''' Return the animal numbers found in the daily report
        '''
        animal_numbers = set()
        for row_number in range(1, self._sheet.nrows):
            animal_number = self._sheet.row_values(row_number)[self.ANIMAL_ID_COL]
            # xls stores all numbers as float, but also handle str type just in case
            if isinstance(animal_number, float) or (isinstance(animal_number, str) and animal_number.isdigit()):
                animal_numbers.add((int(animal_number)))

        return animal_numbers

    def output_results(self, persons_data, foster_parents, animal_details, filtered_animals, csv_filename, dog_mode):
        ''' Output the results to a new csv document
        '''
        print_success('Writing results to {}...'.format(csv_filename))
        csv_rows = []

        csv_rows.append([])
        csv_rows[-1].append('Kitten-Scraper Notes')
        csv_rows[-1].append('Loss Rate')
        csv_rows[-1].append('Name')
        csv_rows[-1].append('E-mail')
        csv_rows[-1].append('Phone')
        csv_rows[-1].append('Person ID')
        csv_rows[-1].append('Foster Experience')
        csv_rows[-1].append('Date {} Received'.format('Kittens' if not dog_mode else 'Canines'))
        csv_rows[-1].append('Details')
        csv_rows[-1].append('Special Animal Message')

        for person_number in sorted(foster_parents, key=lambda p: persons_data[p]['notes']):
            person_data = persons_data[person_number]
            name = person_data['full_name']
            prev_animals_fostered = person_data['prev_animals_fostered']
            report_notes = person_data['notes']
            loss_rate = round(person_data['loss_rate'])

            animal_details_string, animal_numbers = self._collect_animals(person_number,
                                                                          foster_parents,
                                                                          animal_details)
            special_message = ''
            for a in animal_numbers:
                msg = animal_details[a]['message']
                if msg:
                    special_message += '{}{}: {}'.format('\r\r' if special_message else '', a, animal_details[a]['message'])

            if prev_animals_fostered is not None:
                foster_experience = 'NEW' if not prev_animals_fostered else '{}'.format(prev_animals_fostered)
            else:
                foster_experience = ''

            # Build phone number(s) string
            #
            cell_number = person_data['cell_phone']
            home_number = person_data['home_phone']
            phone = ''
            if len(cell_number) >= 10: # ignore incomplete phone numbers
                phone = '(C) {}'.format(cell_number)

            if len(home_number) >= 10: # ignore incomplete phone numbers
                phone += '{}(H) {}'.format('\r' if phone else '', home_number)

            # Build email(s) string
            #
            email = ''
            for e in person_data['emails']:
                email += '{}{}'.format('\r' if email else '', e)

            # Since we're processing a "daily report" I can assume all animals in this group went into foster on the
            # same date
            #
            date_received = animal_details[animal_numbers[0]]['status_date']

            # Explicitly wrapping numbers/datestr with ="{}" to avoid Excel auto-formatting issues
            #
            csv_rows.append([])
            csv_rows[-1].append('"{}"'.format(report_notes))
            csv_rows[-1].append('"{}%"'.format(loss_rate))
            csv_rows[-1].append('"{}"'.format(name))
            csv_rows[-1].append('"{}"'.format(email))
            csv_rows[-1].append('"{}"'.format(phone))
            csv_rows[-1].append('="{}"'.format(person_number))
            csv_rows[-1].append('"{}"'.format(foster_experience))
            csv_rows[-1].append('="{}"'.format(date_received))
            csv_rows[-1].append('"{}"'.format(animal_details_string))
            csv_rows[-1].append('"{}"'.format(special_message))

            print('{} (Experience: {}, Loss Rate: {}%) {}{}{}'.format(name, foster_experience, loss_rate, ConsoleFormat.GREEN, report_notes.replace('\r', ', '), ConsoleFormat.END))

        with open(csv_filename, 'w') as outfile:
            for row in csv_rows:
                outfile.write(','.join(row))
                outfile.write('\n')

            if not len(foster_parents):
                outfile.write('*** None of the animals in this report are currently in foster\n')
                print_warn('None of the animals in this report are currently in foster. Nothing to do!')

            if len(filtered_animals):
                outfile.write('\n\n\n*** Animals not in foster\n')
                print_warn('\nAnimals not in foster')
                for a in filtered_animals:
                    outfile.write('{} - {}\n'.format(a, animal_details[a]['status']))
                    print('{} - {}'.format(a, animal_details[a]['status']))

    def _xlsfloat_as_datetime(self, xlsfloat, workbook_datemode):
        ''' Convert Excel float date type to datetime
        '''
        if not xlsfloat:
            return None

        return datetime(*xlrd.xldate_as_tuple(xlsfloat, workbook_datemode))

    def _copy_row_as_text(self, row_number):
        ''' Output is written to csv so we need to stringify all types (dates in particular).
            Explicitly wrapping numbers/datestr with ="{}" to avoid Excel auto-formatting issues.
        '''
        values = []
        for col_number in range(0, len(self._sheet.row_values(row_number))):
            values.append(self._cell_to_string(row_number, col_number))

        return values

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

    def _pretty_print_animal_age(self, age_string):
        ''' Expecting an age string in the format '%d years %d months %d weeks'
        '''
        pretty_age_string = ''
        animal_type = ''
        try:
            # For the sake of brevity in the spreadsheet, I'll shorten the age string when I can.
            # For example, if an animal is > 1 year old, there's no huge need to include months and weeks.
            # I'll also generate an animal_type string (kitten/cat/puppy/dog) based on age, since daily reports
            # occasionally fail to include this information.
            #
            (years, months, weeks) = re.search(r'(\d+) years (\d+) months (\d+) weeks', age_string).groups()
            if int(years) > 0:
                pretty_age_string = '{} years'.format(years)
                animal_type = self.ADULT_ANIMAL_TYPE
            elif int(months) >= 3:
                pretty_age_string = '{} months'.format(months)
                animal_type = self.YOUNG_ANIMAL_TYPE if int(months) <= 6 else self.ADULT_ANIMAL_TYPE
            else:
                pretty_age_string = '{} weeks'.format(int(weeks) + int(months) * 4)
                animal_type = self.YOUNG_ANIMAL_TYPE
        except:
            animal_type = self.YOUNG_ANIMAL_TYPE # I have no idea, but self.YOUNG_ANIMAL_TYPE is a safe bet
            pretty_age_string = 'Unknown Age'

        return pretty_age_string, animal_type

    def _collect_animals(self, person_number, foster_parents, animal_details):
        ''' Collect/format additional details for each animal assigned to person_number
        '''
        animal_types = []
        animals_by_type = {}
        animal_numbers = []
        animal_ages = {}
        last_animal_type = ''
        for row_number in range(1, self._sheet.nrows): # ignore header
            a_number = int(self._sheet.row_values(row_number)[self.ANIMAL_ID_COL] or 0)
            if not a_number:
                continue

            a_type = self._sheet.row_values(row_number)[self.ANIMAL_TYPE_COL]
            a_age = self._sheet.row_values(row_number)[self.ANIMAL_AGE_COL]
            p_number = next((p for p in foster_parents if a_number in foster_parents[p]), None)

            if not a_type:
                a_type = last_animal_type
            else:
                a_type = utf8(a_type)
                last_animal_type = a_type

            if person_number == p_number and a_number not in animal_numbers:
                animal_types.append(a_type)
                animal_numbers.append(a_number)

                # NOTE: Making an assumption here that the same animal types will be of the same age, or at least
                # close enough to be grouped together. Therefore choosing one age to share won't be much of an issue.
                #
                if a_age:
                    animal_ages[a_type] = a_age

                animals_by_type.setdefault(a_type, []).append(a_number)

        animal_counts_by_type = {}
        for animal in set(animal_types):
            animal_counts_by_type[animal] = animal_types.count(animal)

        # Pretty-print results
        #
        details_str = ''
        for animal_type in animal_counts_by_type:
            _age, _type = self._pretty_print_animal_age(animal_ages[animal_type] if animal_type in animal_ages else '')

            _details = ''
            for a in animals_by_type[animal_type]:
                _sn = animal_details[a]['sn'] if a in animal_details else 'Unknown'
                _status = animal_details[a]['status'] if a in animal_details else 'Status: Unknown'
                _name = animal_details[a]['name'] if a in animal_details else 'No Name'
                _breed = animal_details[a]['breed'] if a in animal_details else 'Unknown Breed'
                _color = animal_details[a]['primary_color'] if a in animal_details else 'Unknown Color'
                _secondary_color = animal_details[a]['secondary_color'] if a in animal_details else None
                if _secondary_color not in [None, 'None']:
                    _color += '/{}'.format(_secondary_color)

                _details += '{}{} {}, S/N {}, {}, {}, {}'.format('\r' if _details else '',
                                                          a,
                                                          _status,
                                                          _sn,
                                                          _name,
                                                          _breed,
                                                          _color)

            details_str += '{}{} {}{} @ {}\r{}'.format('\r' if details_str else '',
                                                        animal_counts_by_type[animal_type],
                                                        _type.lower(),
                                                        's' if animal_counts_by_type[animal_type] > 1 else '',
                                                        _age,
                                                        _details)
        return details_str, animal_numbers
