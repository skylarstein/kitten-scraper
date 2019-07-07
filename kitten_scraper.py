from __future__ import print_function
from __init__ import __version__
import os
import re
import sys
import time
import yaml
from argparse import ArgumentParser
from datetime import datetime
from google_sheets_reader import GoogleSheetsReader
from kitten_report_reader import KittenReportReader
from kitten_utils import *
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class KittenScraper(object):
    def load_config_file(self):
        ''' A config.yaml configuration file is expected to be in the same directory as this script
        '''
        try:
            config_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'config.yaml')
            config = yaml.load(open(config_file, 'r'))
            self._username = config['username']
            self._password = config['password']
            self._mentors_spreadsheet_key = config['mentors_spreadsheet_key']

        except yaml.YAMLError as err:
            print_err('ERROR: Unable to parse configuration file: {}, {}'.format(config_file, err))
            return False

        except IOError as err:
            print_err('ERROR: Unable to read configuration file: {}, {}'.format(config_file, err))
            return False

        except KeyError as err:
            print_err('ERROR: Missing value in configuration file: {}, {}'.format(config_file, err))
            return False

        return True

    def read_config_yaml(self, config_yaml):
        ''' The mentors spreadsheet contains additional configuration data. This makes it easier to manage dynamic
            configuration data vs rollout of config.yaml updates.
        '''
        try:
            config = yaml.load(config_yaml)
            self._login_url = config['login_url']
            self._search_url = config['search_url']
            self._animal_url = config['animal_url']
            self._medical_details_url = config['medical_details_url']
            self._list_animals_url = config['list_animals_url']
            self._do_not_assign_mentor = config['do_not_assign_mentor'] if 'do_not_assign_mentor' in config else []
            self._mentors = config['mentors'] if 'mentors' in config else []

        except yaml.YAMLError as err:
            print_err('ERROR: Unable to read config_yaml: {}'.format(err))
            return False

        except KeyError as err:
            print_err('ERROR: Missing value in config_yaml: {}'.format(err))
            return False

        return True

    def start_browser(self, show_browser):
        ''' Instantiate the browser, configure options as needed
        '''
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/605.1.15 '
                                    '(KHTML, like Gecko) Version/12.0.3 Safari/605.1.15')
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

        self._driver = webdriver.Chrome(chromedriver_path, options = chrome_options)
        self._driver.set_page_load_timeout(60)

    def exit_browser(self):
        ''' Close and exit the browser
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
        except NoSuchElementException:
            print_err('ERROR: Unable to load the login page. Please check your connection.')
            return False

        try:
            self._driver.find_element_by_id("txt_username").send_keys(self._username)
            self._driver.find_element_by_id("txt_password").send_keys(self._password)
            self._driver.find_element_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_btn_login').click()
            self._driver.find_element_by_id('Continue').click()

        except NoSuchElementException:
            print_err('ERROR: Unable to login. Please check your username/password.')
            return False

        return True

    def get_animal_details(self, animal_numbers):
        foster_parents = {}
        animal_details = {}
        filtered_animals = set()
        for a in animal_numbers:
            print('Looking up animal {}... '.format(a), end='')
            sys.stdout.flush()
            self._driver.get(self._animal_url.format(a))
            try:
                # Dismiss alert (if found)
                #
                Alert(self._driver).dismiss()
            except NoAlertPresentException:
                pass

            try:
                # Wait for lazy-loaded content
                #
                WebDriverWait(self._driver, 10).until(EC.presence_of_element_located((By.ID, 'submitbtn2')))
            except:
                raise Exception('Timeout while waiting for content on search page!')

            # Get Special Message text (if it exists)
            #
            special_msg = utf8(self._get_text_by_id('specialMessagesDialog'))

            if special_msg:
                # Remove text we don't care about
                #
                special_msg = re.sub(r'(?i)This is a special message. If you would like to delete it then clear the Special Message box in the General Details section of this page.', '', special_msg).strip()

                # Remove empty lines
                #
                special_msg = os.linesep.join([s for s in special_msg.splitlines() if s])

            animal_details[a] = lambda: None # emulate SimpleNamespace for Python 2.7
            animal_details[a].message = special_msg
            status = self._get_selection_by_id('status')
            sub_status = self._get_selection_by_id('subStatus')
            animal_details[a].status = '{}{}{}'.format(status, ' - ' if sub_status else '', sub_status)

            try:
                animal_details[a].status_date = datetime.strptime(self._get_attr_by_id('statusdate'), '%m/%d/%Y').strftime('%-d-%b-%Y')
            except ValueError:
                animal_details[a].status_date = 'Unknown'

            # If this animal is currently in foster, get the responsible person (foster parent)
            #
            if status.lower().find('in foster') >= 0 and status.lower().find('unassisted death') < 0:
                try:
                    p = int(self._get_attr_by_xpath('href', '//*[@id="Table17"]/tbody/tr[1]/td[2]/a').split('personid=')[1])
                    foster_parents.setdefault(p, []).append(a)
                except:
                    print_err('Failed to find foster parent for animal {}, please check report'.format(a))
            else:
                filtered_animals.add(a)

            # Load spay/neuter status from the medical details page
            #
            try:
                self._driver.get(self._medical_details_url.format(a))
                animal_details[a].sn = utf8(self._get_attr_by_xpath('innerText', '/html/body/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[4]'))
            except:
                print_err('Failed to read spay/neuter status for animal {}'.format(a))
                animal_details[a].sn = 'Unknown'

            print('{}'.format(animal_details[a].status))

        return foster_parents, animal_details, filtered_animals

    def get_person_data(self, person_number, google_sheets_reader):
        ''' Load the given person number, return details and contact information
        '''
        print('Looking up person {}... '.format(person_number), end='')
        sys.stdout.flush()

        self._driver.get(self._search_url)
        self._driver.find_element_by_id('userid').send_keys(str(person_number))
        self._driver.find_element_by_id('userid').send_keys(webdriver.common.keys.Keys.RETURN)

        first_name     = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtFirstName')
        last_name      = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtLastName')
        preferred_name = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtPreferredName')
        home_phone     = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_homePhone_txtPhone3')
        cell_phone     = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_mobilePhone_txtPhone3')
        email1         = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[1]/td[1]')
        email2         = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[2]/td[1]')
        email3         = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[3]/td[1]')
        email4         = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[4]/td[1]')

        emails = set()
        for email in (email for email in [email1, email2, email3, email4] if email): # add non-empties to set
            emails.add(email.lower())

        prev_animals_fostered, euthanized_count, unassisted_death_count = self._prev_animals_fostered(person_number)

        full_name = preferred_name if preferred_name else first_name if first_name else ''
        full_name += ' ' if len(full_name) else ''
        full_name += last_name if last_name else ''

        notes = ''

        # 'do_not_assign_mentor' list may included person number (as number) or email (as string)
        #
        if person_number in self._do_not_assign_mentor:
            notes = '*** Do not assign mentor (Staff)'
        else:
            do_not_assign_strings = (s for s in self._do_not_assign_mentor if isinstance(s, str))
            for dna in do_not_assign_strings:
                if [s for s in emails if dna.lower() in s]:
                    notes = '*** Do not assign mentor (Staff)'
                    break

        if person_number in self._mentors:
            notes += '{}*** {} is a mentor'.format('\r' if len(notes) else '', full_name)

        match_strings = emails.union([full_name])
        matching_sheets = google_sheets_reader.find_matches_in_feline_foster_spreadsheet(match_strings)
        if matching_sheets:
            notes += '{}*** Found matching mentor(s): {}'.format('\r' if len(notes) else '', ', '.join([str(s) for s in matching_sheets]))

        loss_rate = 0.0
        if prev_animals_fostered > 0:
            loss_rate = 100.0 * (euthanized_count + unassisted_death_count) / prev_animals_fostered

        print('{} {}'.format(first_name, last_name))
        return {
            'first_name'             : first_name,
            'last_name'              : last_name,
            'preferred_name'         : preferred_name,
            'full_name'              : full_name,
            'home_phone'             : home_phone,
            'cell_phone'             : cell_phone,
            'emails'                 : emails,
            'prev_animals_fostered'  : prev_animals_fostered,
            'euthanized_count'       : euthanized_count,
            'unassisted_death_count' : unassisted_death_count,
            'loss_rate'              : loss_rate,
            'notes'                  : notes
        }

    def _prev_animals_fostered(self, person_number):
        ''' Determine the total number of felines this person previously fostered. This is a useful metric for
            experience level.

            Load the list of all animals this person has been responsible for, page by page until we have no more pages.
        '''
        page_number = 1
        previous_feline_foster_count = 0
        euthanized_count = 0
        unassisted_death_count = 0

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
                            if animal_type.lower().find('cat') >= 0 or animal_type.lower().find('kitten') >= 0:
                                if animal_status.lower().find('in foster') < 0:
                                    previous_feline_foster_count += 1

                                if animal_status.lower().find('euthanized') >= 0:
                                    euthanized_count += 1
                                elif animal_status.lower().find('unassisted death') >= 0:
                                    unassisted_death_count += 1

                    elif num_cols == 1 and cols[0].text == 'Fostered':
                        foster_tr_active = True

                    else:
                        foster_tr_active = False

            except NoSuchElementException:
                break

            page_number = page_number + 1
        return previous_feline_foster_count, euthanized_count, unassisted_death_count

    def _get_text_by_id(self, element_id):
        try:
            return self._driver.find_element_by_id(element_id).text
        except:
            return ''

    def _get_attr_by_id(self, element_id):
        try:
            return self._driver.find_element_by_id(element_id).get_attribute('value')
        except:
            return ''

    def _get_attr_by_xpath(self, attr, element_xpath):
        try:
            return self._driver.find_element_by_xpath(element_xpath).get_attribute(attr)
        except:
            return ''

    def _get_checked_by_id(self, element_id):
        try:
            attr = self._driver.find_element_by_id(element_id).get_attribute('checked')
            return True if attr else False
        except:
            return False

    def _get_selection_by_id(self, element_id):
        try:
            select_element = Select(self._driver.find_element_by_id(element_id))
            return select_element.first_selected_option.text
        except:
            return ''

if __name__ == "__main__":
    print('Welcome to KittenScraper {}'.format(__version__))
    start_time = time.time()

    arg_parser = ArgumentParser()
    arg_parser.add_argument('-i', '--input', help = 'daily kitten report (xls)', required = False)
    arg_parser.add_argument('-o', '--output', help = 'output file (csv)', required = False)
    arg_parser.add_argument('--show_browser', help = 'show the browser window while working', required = False, action = 'store_true')
    args = arg_parser.parse_args()

    if not args.input or not args.output:
        arg_parser.print_help()
        sys.exit(0)

    # Load config.yaml
    #
    kitten_scraper = KittenScraper()
    if not kitten_scraper.load_config_file():
        sys.exit()

    # Load the "daily report" xls
    #
    kitten_report_reader = KittenReportReader()
    if not kitten_report_reader.open_xls(args.input):
        sys.exit()

    # Assure the output path exists
    #
    output_path = os.path.dirname(args.output)
    if output_path and not os.path.exists(output_path):
        os.makedirs(output_path)

    # Load the Feline Mentors spreadsheet
    #
    google_sheets_reader = GoogleSheetsReader()
    config_yaml = google_sheets_reader.load_mentors_spreadsheet(sheets_key = kitten_scraper._mentors_spreadsheet_key)
    if not config_yaml:
        sys.exit()

    if not kitten_scraper.read_config_yaml(config_yaml):
        sys.exit()

    # Process the daily report
    #
    animal_numbers = kitten_report_reader.get_animal_numbers()
    print('Found {} animal numbers: {}'.format(len(animal_numbers), ', '.join([str(a) for a in animal_numbers])))

    kitten_scraper.start_browser(args.show_browser)
    if not kitten_scraper.login():
        sys.exit()

    # Query details for each animal (current foster parent, foster status, etc)
    #
    foster_parents, animal_details, filtered_animals = kitten_scraper.get_animal_details(animal_numbers)

    for p in foster_parents:
        print('Animals for foster parent {} = {}'.format(p, foster_parents[p]))

    # Query foster parent details (person number -> name, contact details, etc)
    #
    persons_data = {}
    for person in foster_parents:
        persons_data[person] = kitten_scraper.get_person_data(person, google_sheets_reader)

    kitten_scraper.exit_browser()

    # Output the combined results to csv
    #
    kitten_report_reader.output_results(persons_data, foster_parents, animal_details, filtered_animals, args.output)

    print('\nKitten foster report completed in {0:.3f} seconds'.format(time.time() - start_time))
