import os
import re
import sys
import time
import xlrd
import yaml
from datetime import datetime
from argparse import ArgumentParser
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By

class KittenScraper():

    def __init__(self):
        self.LOGIN_URL = 'http://192.168.100.27/login.aspx?aspInitiated=1.1'
        self.SEARCH_URL = 'http://192.168.100.27/main.asp'

        # The "List All Animals" pagesize=20 query string may lead you to believe you can request more than
        # 20 records per page, but alas, it always returns 20 regardless of this value
        #
        self.LIST_ANIMALS_URL_FORMAT = 'http://192.168.100.27/person/listAll.asp?tpage={}&pagesize=20&task=view&recnum={}'

    def exit(self):
        ''' Close and exit the browser instance
        '''
        self.driver.close()
        self.driver.quit()

    def load_configuration(self):
        ''' A config.yaml configuration file is expected to be in the same directory as this script
        '''
        try:
            config_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'config.yaml')
            config = yaml.load(open(config_file, 'r'))
            self.username = config['username']
            self.password = config['password']
            return True

        except yaml.YAMLError as err:
            print('ERROR: Unable to parse configuration file: {}, {}'.format(config_file, err))

        except IOError as err:
            print('ERROR: Unable to read configuration file: {}, {}'.format(config_file, err))

        return False

    def start_browser(self, show_browser):
        ''' Instantiate the browser, configure options as needed
        '''
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.117 Safari/537.36')
        if not show_browser:
            chrome_options.add_argument("--headless")

        chromedriver_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bin/mac64/chromedriver')
        self.driver = webdriver.Chrome(chromedriver_path, chrome_options = chrome_options)

    def login(self):
        ''' Load the login page, enter credentials, submit
        '''
        print('Logging in...')
        self.driver.get(self.LOGIN_URL)

        self.driver.find_element_by_id("txt_username").send_keys(self.username)
        self.driver.find_element_by_id("txt_password").send_keys(self.password)
        self.driver.find_element_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_btn_login').click()
        self.driver.find_element_by_id('Continue').click()

    def get_person_data(self, person_number):
        ''' Search for the given person number, return details and contact information
        '''
        print('Looking up person number {}...'.format(person_number))

        self.driver.get(self.SEARCH_URL)
        self.driver.find_element_by_id("userid").send_keys(person_number)
        self.driver.find_element_by_id("userid").send_keys(webdriver.common.keys.Keys.RETURN)

        first_name            = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtFirstName')
        last_name             = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtLastName')
        preferred_name        = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtPreferredName')
        home_phone            = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_homePhone_txtPhone3')
        cell_phone            = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_mobilePhone_txtPhone3')		
        primary_email         = self.get_text_by_xpath('//*[@id="emailTable"]/tbody/tr[1]/td[1]')
        secondary_email       = self.get_text_by_xpath('//*[@id="emailTable"]/tbody/tr[2]/td[1]')
        prev_animals_fostered = self.prev_animals_fostered(person_number)

        return {
            'first_name'            : first_name,
            'last_name'             : last_name,
            'preferred_name'        : preferred_name,
            'home_phone'            : home_phone,
            'cell_phone'            : cell_phone,
            'primary_email'         : primary_email,
            'secondary_email'       : secondary_email,
            'prev_animals_fostered' : prev_animals_fostered
        }

    def prev_animals_fostered(self, person_number):
        ''' Determine the total number of felines this person previously fostered. This is used
            to determine feline foster experience level.

            Load the list of all animals this person has been responsible for, page by page until
            we have no more pages.
        '''
        page_number = 1
        previous_feline_foster_count = 0
        while True:
            self.driver.get(self.LIST_ANIMALS_URL_FORMAT.format(page_number, person_number))
            try:
                table = self.driver.find_element_by_id('Table3')
                rows = table.find_elements(By.TAG_NAME, 'tr')
                foster_tr_active = False

                for row in rows:
                    cols = row.find_elements(By.TAG_NAME, 'td')
                    num_cols = len(cols)

                    if num_cols == 10: 
                        if foster_tr_active:
                            #print ','.join(col.text for col in cols)
                            animal_type = cols[5].text
                            animal_status = cols[2].text
                            if (animal_type == 'Cat' or animal_type == 'Kitten') and animal_status != 'In Foster':
                                previous_feline_foster_count += 1

                    elif num_cols == 1 and cols[0].text == 'Fostered':
                        foster_tr_active = True

                    else:
                        foster_tr_active = False

            except NoSuchElementException:
                break

            page_number = page_number + 1
        return previous_feline_foster_count

    def get_text_by_id(self, element_id):
        ''' Quick helper function to get attribute text with proper error handling
        '''
        try:
            return self.driver.find_element_by_id(element_id).get_attribute('value')
        except:
            return ''

    def get_text_by_xpath(self, element_xpath):
        ''' Quick helper function to get attribute text with proper error handling
        '''
        try:
            return self.driver.find_element_by_xpath(element_xpath).get_attribute('innerText')
        except:
            return ''

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

            print('Loaded report {}'.format(xls_filename))
            return True

        except IOError as err:
            print('ERROR: Unable to read xls file: {}, {}'.format(xls_filename, err))

        except xlrd.XLRDError as err:
            print('ERROR: Unable to read xls file: {}'.format(err.message))

        return False

    def get_person_numbers(self):
        ''' Return a set of unique person numbers found in the daily report
        '''
        persons = set()
        for row_number in range(1, self.sheet.nrows):
            # If a person number has no associated animal number, this is due to a bug in the
            # report which includes a mostly empty row for an animal's previous foster parent.
            # Ignore these cases.
            #
            animal_number = self.sheet.row_values(row_number)[2]
            person_number = self.sheet.row_values(row_number)[5]

            if isinstance(animal_number, float): # xls stores all numbers as float
                persons.add(str(int(person_number)))

        return persons

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
            pass

        return result

    def count_animals(self, person_number):
        ''' Count the number and age of each animal type assigned to this person number
        '''
        animals = []
        animals_age = {}
        last_animal_type = ''
        for row_number in range(1, self.sheet.nrows): # ignore header
            a_type = self.sheet.row_values(row_number)[1]
            a_age = self.sheet.row_values(row_number)[4]
            p_number = self.sheet.row_values(row_number)[5]

            if not a_type:
                a_type = last_animal_type
            else:
                last_animal_type = a_type

            if person_number == p_number:
                animals.append(a_type)
                # WARNING: Making an assumption here that same animal types will be of the same age
                # (or at least close enough in grouped litters) so that choosing one age to share won't
                # be much of an issue
                #
                animals_age[a_type] = a_age

        # We now have a list of all animal types. Next, create a set with total counts per type.
        #
        animal_counts = {}
        for animal in set(animals):
            animal_counts[animal] = animals.count(animal)

        # Pretty-print results
        #
        result_str = ''
        for animal in animal_counts:
            if result_str:
                result_str += '\r'
            age = self.pretty_print_animal_age(animals_age[animal])
            result_str += '{} {}{} @ {}'.format(animal_counts[animal], animal, 's' if animal_counts[animal] > 1 else '', age)

        return result_str.lower()

    def output_results(self, persons_data, csv_filename):
        ''' Combine the newly gathered person data with the daily report, output results
            to a new csv document.
        '''
        print('Writing results to {}...'.format(csv_filename))

        new_rows = []

        # First include the original column headers, then add columns for our new data
        #
        new_rows.append(self.sheet.row_values(0))
        new_rows[-1].append('Name')
        new_rows[-1].append('E-mail')
        new_rows[-1].append('Phone')
        new_rows[-1].append('Foster Experience')
        new_rows[-1].append('Date Kittens Received')	
        new_rows[-1].append('Quantity')

        for row_number in range(1, self.sheet.nrows): # ignore header
            animal_type = self.sheet.row_values(row_number)[1]
            animal_number = self.sheet.row_values(row_number)[2]
            person_number = self.sheet.row_values(row_number)[5]
            status_datetime = self.xlsfloat_as_datetime(self.sheet.row_values(row_number)[0], self.workbook.datemode)

            # If there is no animal number in this row, skip the row
            #
            if not animal_number:
                continue

            # Include original column data as text since we're building a CSV document
            #
            new_rows.append(self.copy_row_as_text(row_number))

            # Only include person details for rows with 'Current Animal Type' populated
            #
            if not animal_type:
                continue

            # Grab the person data from the associated person number
            #
            person_number_str = str(int(person_number))
            person_data = persons_data[person_number_str] if person_number_str in persons_data else {}

            # Build full name
            #
            name = person_data['preferred_name'] if 'preferred_name' in person_data else ''
            if not len(name):
                name = person_data['first_name'] if 'first_name' in person_data else ''

            name += ' '
            name += person_data['last_name'] if 'last_name' in person_data else ''

            # Build phone number(s)
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

            # Build email(s)
            #
            email = person_data['primary_email'] if 'primary_email' in person_data else ''
            secondary_email = person_data['secondary_email'] if 'secondary_email' in person_data else ''

            if len(secondary_email):
                if len(email):
                    email += '\r'
                email += secondary_email

            animal_quantity = self.count_animals(person_number)
            prev_animals_fostered = person_data['prev_animals_fostered']
            foster_experience = 'NEW' if not prev_animals_fostered else '{} previous'.format(prev_animals_fostered)

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

        with open(csv_filename, 'w') as outfile:
            for row in new_rows:
                outfile.write(','.join(row))
                outfile.write('\n')

if __name__ == "__main__":
    start_time = time.time()

    arg_parser = ArgumentParser()
    arg_parser.add_argument('-i', '--input', help = 'daily kitten report (xls)', required = False)
    arg_parser.add_argument('-o', '--output', help = 'output file (csv)', required = False)
    arg_parser.add_argument('--show_browser', help = 'show the browser window while working', required = False, action = 'store_true')
    args = arg_parser.parse_args()

    if not args.input or not args.output:
        arg_parser.print_help()
        sys.exit(0)

    # Load me up some kittens and foster parent numbers from the xls report
    #
    kitten_reader = KittenReportReader()
    if not kitten_reader.open_xls(args.input):
        sys.exit()

    persons = kitten_reader.get_person_numbers()
    print('Found foster parent numbers: {}'.format(persons))

    # Log in and query foster parent details (person number -> name and contact details)
    #
    kitten_scraper = KittenScraper()
    if not kitten_scraper.load_configuration():
        sys.exit()

    kitten_scraper.start_browser(args.show_browser)
    kitten_scraper.login()

    persons_data = {}
    for person in persons:
        persons_data[person] = kitten_scraper.get_person_data(person)

    kitten_scraper.exit()

    # Output the combined results to csv
    #
    kitten_reader.output_results(persons_data, args.output)

    print('\nKitten scraping completed in {0:.3f} seconds'.format(time.time() - start_time))
