import os
import sys
import time
import yaml
from argparse import ArgumentParser
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from google_sheets_reader import GoogleSheetsReader
from kitten_report_reader import KittenReportReader
from kitten_utils import *

class KittenScraper:
    def __init__(self):
        pass

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
            self.login_url = config['login_url']
            self.search_url = config['search_url']
            self.list_animals_url = config['list_animals_url']
            self.do_not_assign_mentor = config['do_not_assign_mentor'] if 'do_not_assign_mentor' in config else []
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

        chromedriver_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bin/mac64/chromedriver')
        self.driver = webdriver.Chrome(chromedriver_path, chrome_options = chrome_options)

    def login(self):
        ''' Load the login page, enter credentials, submit
        '''
        print('Logging in...')
        self.driver.get(self.login_url)

        self.driver.find_element_by_id("txt_username").send_keys(self.username)
        self.driver.find_element_by_id("txt_password").send_keys(self.password)
        self.driver.find_element_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_btn_login').click()
        self.driver.find_element_by_id('Continue').click()

    def get_person_data(self, person_number, google_sheets_reader):
        ''' Search for the given person number, return details and contact information
        '''
        print('Looking up person number {}...'.format(person_number))

        self.driver.get(self.search_url)
        self.driver.find_element_by_id("userid").send_keys(str(person_number))
        self.driver.find_element_by_id("userid").send_keys(webdriver.common.keys.Keys.RETURN)

        first_name            = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtFirstName')
        last_name             = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtLastName')
        preferred_name        = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtPreferredName')
        home_phone            = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_homePhone_txtPhone3')
        cell_phone            = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_mobilePhone_txtPhone3')		
        primary_email         = self.get_text_by_xpath('//*[@id="emailTable"]/tbody/tr[1]/td[1]')
        secondary_email       = self.get_text_by_xpath('//*[@id="emailTable"]/tbody/tr[2]/td[1]')
        prev_animals_fostered = self.prev_animals_fostered(person_number)

        full_name = preferred_name if preferred_name else first_name if first_name else ''
        full_name += ' ' if len(full_name) else ''
        full_name += last_name if last_name else ''

        notes = '*** Do Not Assign Mentor' if person_number in self.do_not_assign_mentor else ''
        matching_sheets = google_sheets_reader.find_matches_in_feline_foster_spreadsheet([full_name, primary_email, secondary_email])
        if matching_sheets:
            notes += '\r' if len(notes) else ''
            notes += '*** Found matching mentor(s): {}'.format(', ' .join([str(s) for s in matching_sheets]))

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

    def prev_animals_fostered(self, person_number):
        ''' Determine the total number of felines this person previously fostered. This is used
            to determine feline foster experience level.

            Load the list of all animals this person has been responsible for, page by page until
            we have no more pages.
        '''
        page_number = 1
        previous_feline_foster_count = 0
        while True:
            self.driver.get(self.list_animals_url.format(page_number, person_number))
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
    kitten_report_reader = KittenReportReader()
    if not kitten_report_reader.open_xls(args.input):
        sys.exit()

    persons = kitten_report_reader.get_person_numbers()
    print('Found foster parent numbers: {}'.format(', '.join([str(person) for person in persons])))

    # Load the Feline Mentors spreadsheet
    #
    google_sheets_reader = GoogleSheetsReader()
    google_sheets_reader.load_mentors_spreadsheet(sheets_key='1XuZqVA7t2yGbKzcnCAsUuMW-W44L3PKreRWlSHZUYLc')
    print('Loaded {} worksheets from mentors spreadsheet'.format(len(google_sheets_reader.sheet_data)))

    # Log in and query foster parent details (person number -> name and contact details)
    #
    kitten_scraper = KittenScraper()
    if not kitten_scraper.load_configuration():
        sys.exit()

    print('Special config (do not assign mentor): {}'.format(', '.join(str(p) for p in kitten_scraper.do_not_assign_mentor)))

    kitten_scraper.start_browser(args.show_browser)
    kitten_scraper.login()

    persons_data = {}
    for person in persons:
        persons_data[person] = kitten_scraper.get_person_data(person, google_sheets_reader)

    kitten_scraper.exit()

    # Output the combined results to csv
    #
    kitten_report_reader.output_results(persons_data, args.output)

    print('\nKitten scraping completed in {0:.3f} seconds'.format(time.time() - start_time))
