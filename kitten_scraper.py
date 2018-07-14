import os
import re
import sys
import time
import yaml
from argparse import ArgumentParser
from google_sheets_reader import GoogleSheetsReader
from kitten_report_reader import KittenReportReader
from kitten_utils import *
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By

class KittenScraper(object):
    def load_configuration(self):
        ''' A config.yaml configuration file is expected to be in the same directory as this script
        '''
        try:
            config_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'config.yaml')
            config = yaml.load(open(config_file, 'r'))

            self._username = config['username']
            self._password = config['password']
            self._login_url = config['login_url']
            self._search_url = config['search_url']
            self._animal_url = config['animal_url']
            self._list_animals_url = config['list_animals_url']
            self._do_not_assign_mentor = config['do_not_assign_mentor'] if 'do_not_assign_mentor' in config else []
            self._mentors = config['mentors'] if 'mentors' in config else []

            return True

        except yaml.YAMLError as err:
            print_err('ERROR: Unable to parse configuration file: {}, {}'.format(config_file, err))

        except IOError as err:
            print_err('ERROR: Unable to read configuration file: {}, {}'.format(config_file, err))

        return False

    def start_browser(self, show_browser):
        ''' Instantiate the browser, configure options as needed
        '''
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.117 Safari/537.36')
        if not show_browser:
            chrome_options.add_argument("--headless")

        # TODO: Consider adding chromedriver-binary or chromedriver_installer to requirements.txt and
        # removing these local copies...
        if sys.platform == 'darwin':
            chromedriver_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bin/mac64/chromedriver')
        elif sys.platform.startswith('linux'):
            chromedriver_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bin/linux64/chromedriver')
        else:
            print_err('Sorry friends, I haven\'t included chromedriver for your platform ({}). Exiting now.'.format(sys.platform))
            sys.exit(0)

        self._driver = webdriver.Chrome(chromedriver_path, chrome_options = chrome_options)

    def exit_browser(self):
        ''' Close and exit the browser instance
        '''
        self._driver.close()
        self._driver.quit()

    def login(self):
        ''' Load the login page, enter credentials, submit
        '''
        print_success('Logging in...')

        try:
            self._driver.set_page_load_timeout(20)
            self._driver.get(self._login_url)
        except TimeoutException:
            print_err('ERROR: Unable to load the login page. Please check your connection.')
            return False

        self._driver.find_element_by_id("txt_username").send_keys(self._username)
        self._driver.find_element_by_id("txt_password").send_keys(self._password)
        self._driver.find_element_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_btn_login').click()
        self._driver.find_element_by_id('Continue').click()
        return True

    def get_animal_details(self, animal_numbers):
        print('Looking up details for each foster animal...')
        foster_parents = {}
        animal_special_messages = {}
        for a in animal_numbers:
            self._driver.get(self._animal_url.format(a))
            try:
                # Dismiss alert (if found)
                Alert(self._driver).dismiss()
            except NoAlertPresentException:
                pass

            try:
                p = int(self._get_attr_by_xpath('href', '//*[@id="Table17"]/tbody/tr[1]/td[2]/a').split('personid=')[1])
                foster_parents.setdefault(p, []).append(a)
            except:
                print_err('Failed to find foster parent for animal {}, please check report'.format(a))

            # Get Special Message text (if it exists)
            special_msg = utf8(self._get_text_by_id('specialMessagesDialog'))

            if special_msg:
                # Remove text we don't care about...
                special_msg = re.sub(r'(?i)This is a special message. If you would like to delete it then clear the Special Message box in the General Details section of this page.', '', special_msg).strip()

                # Remove empty lines
                special_msg = os.linesep.join([s for s in special_msg.splitlines() if s])

            animal_special_messages[a] = special_msg
 
        return foster_parents, animal_special_messages 

    def get_person_data(self, person_number, google_sheets_reader):
        ''' Search for the given person number, return details and contact information
        '''
        print('Looking up person number {}...'.format(person_number))

        self._driver.get(self._search_url)
        self._driver.find_element_by_id('userid').send_keys(str(person_number))
        self._driver.find_element_by_id('userid').send_keys(webdriver.common.keys.Keys.RETURN)

        first_name            = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtFirstName')
        last_name             = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtLastName')
        preferred_name        = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtPreferredName')
        home_phone            = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_homePhone_txtPhone3')
        cell_phone            = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_mobilePhone_txtPhone3')		
        primary_email         = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[1]/td[1]')
        secondary_email       = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[2]/td[1]')
        prev_animals_fostered = self._prev_animals_fostered(person_number)

        full_name = preferred_name if preferred_name else first_name if first_name else ''
        full_name += ' ' if len(full_name) else ''
        full_name += last_name if last_name else ''

        notes = '*** Do not assign mentor (HSSV Staff)' if person_number in self._do_not_assign_mentor else ''
        if person_number in self._mentors:
            notes += '{}*** {} is a mentor'.format('\r' if len(notes) else '', full_name)

        matching_sheets = google_sheets_reader.find_matches_in_feline_foster_spreadsheet([full_name, primary_email, secondary_email])
        if matching_sheets:
            notes += '{}*** Found matching mentor{}: {}'.format('\r' if len(notes) else '',
                                                                's' if len(matching_sheets) > 1 else '',
                                                                ', '.join([str(s) for s in matching_sheets]))
        return {
            'first_name'            : first_name,
            'last_name'             : last_name,
            'preferred_name'        : preferred_name,
            'full_name'             : full_name,
            'home_phone'            : home_phone,
            'cell_phone'            : cell_phone,
            'primary_email'         : primary_email,
            'secondary_email'       : secondary_email,
            'prev_animals_fostered' : prev_animals_fostered,
            'notes'                 : notes
        }

    def _prev_animals_fostered(self, person_number):
        ''' Determine the total number of felines this person previously fostered. This is used
            to determine feline foster experience level.

            Load the list of all animals this person has been responsible for, page by page until
            we have no more pages.
        '''
        page_number = 1
        previous_feline_foster_count = 0
        while True:
            self._driver.get(self._list_animals_url.format(page_number, person_number))
            try:
                table = self._driver.find_element_by_id('Table3')
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

    def _get_text_by_id(self, element_id):
        ''' Quick helper function to get attribute text with proper error handling
        '''
        try:
            return self._driver.find_element_by_id(element_id).text
        except:
            return ''

    def _get_attr_by_id(self, element_id):
        ''' Quick helper function to get attribute text with proper error handling
        '''
        try:
            return self._driver.find_element_by_id(element_id).get_attribute('value')
        except:
            return ''

    def _get_attr_by_xpath(self, attr, element_xpath):
        ''' Quick helper function to get attribute text with proper error handling
        '''
        try:
            return self._driver.find_element_by_xpath(element_xpath).get_attribute(attr)
        except:
            return ''

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

    kitten_scraper = KittenScraper()
    if not kitten_scraper.load_configuration():
        sys.exit()

    kitten_report_reader = KittenReportReader()
    if not kitten_report_reader.open_xls(args.input):
        sys.exit()

    # Assure output path exists
    #
    output_path = os.path.dirname(args.output)
    if output_path and not os.path.exists(output_path):
        os.makedirs(output_path)

    # Load the daily kitten report
    #
    animal_numbers = kitten_report_reader.get_animal_numbers()
    print('Found {} animal numbers: {}'.format(len(animal_numbers), ', '.join([str(a) for a in animal_numbers])))

    kitten_scraper.start_browser(args.show_browser)
    if not kitten_scraper.login():
        sys.exit()

    # If an animal has had multiple foster parents, the order of associated person IDs may vary per report.
    # The topmost ID may be the current foster parent, or it may be a previous foster parent. This means we
    # need to explicitly look up the current foster parent ID for every kitten number.
    #
    correct_foster_parents, animal_special_messages  = kitten_scraper.get_animal_details(animal_numbers)

    for p in correct_foster_parents:
        print('Animals for foster parent {} = {}'.format(p, correct_foster_parents[p]))

    # Load the Feline Mentors spreadsheet
    #
    google_sheets_reader = GoogleSheetsReader()
    google_sheets_reader.load_mentors_spreadsheet(sheets_key='1XuZqVA7t2yGbKzcnCAsUuMW-W44L3PKreRWlSHZUYLc')

    # Log in and query foster parent details (person number -> name, contact details, etc)
    #
    persons_data = {}
    for person in correct_foster_parents:
        persons_data[person] = kitten_scraper.get_person_data(person, google_sheets_reader)

    kitten_scraper.exit_browser()

    # Output the combined results to csv
    #
    kitten_report_reader.output_results(persons_data, correct_foster_parents, animal_special_messages, args.output)

    print('\nKitten foster report completed in {0:.3f} seconds'.format(time.time() - start_time))
