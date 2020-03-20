from __init__ import __version__
import os
import re
import sys
import time
import yaml
from argparse import ArgumentParser
from datetime import datetime
from box_sheet_reader import BoxSheetReader
from google_sheet_reader import GoogleSheetReader
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
    def run(self):
        print('Welcome to KittenScraper {}'.format(__version__))

        arg_parser = ArgumentParser()
        arg_parser.add_argument('-i', '--input', help = 'specify the daily kitten report (xls)', required = False)
        arg_parser.add_argument('-o', '--output', help = 'specify an output file (csv)', required = False)
        arg_parser.add_argument('-s', '--status', help = 'save current mentee status to the given file (txt)', required = False)
        arg_parser.add_argument('-b', '--show_browser', help = 'show the browser window (generally for debugging)', required = False, action = 'store_true')
        arg_parser.add_argument('-d', '--dog_mode', help = 'enable dog mode', required = False, action = 'store_true')
        args = arg_parser.parse_args()

        if not (args.input and args.output) and not args.status:
            arg_parser.print_help()
            sys.exit(0)

        # Load config.yaml
        #
        if not self._load_config_file():
            sys.exit()

        if args.dog_mode:
            self._dog_mode = True

        if self._dog_mode:
            print_warn('** Canine Mode is Active **')

        # Load the Feline Foster Mentors spreadsheet
        #
        if self._google_spreadsheet_key and self._google_client_secret:
            self.mentor_sheet_reader = GoogleSheetReader()
            self._additional_config_yaml = self.mentor_sheet_reader.load_mentors_spreadsheet({
                'google_spreadsheet_key' : self._google_spreadsheet_key,
                'google_client_secret' : self._google_client_secret})

        elif self._box_user_id and self._box_file_id and self._box_jwt:
            self.mentor_sheet_reader = BoxSheetReader()
            self._additional_config_yaml = self.mentor_sheet_reader.load_mentors_spreadsheet({
                'box_user_id' : self._box_user_id,
                'box_file_id' : self._box_file_id,
                'box_jwt' : self._box_jwt})

        else:
            print_err('ERROR: Incorrect mentor spreadsheet configuration, please check config.yaml')
            sys.exit()

        # Process additional config data from the mentors spreadsheet
        #
        if not self._read_additional_config_yaml(self._additional_config_yaml):
            sys.exit()

        # Start the browser, log in
        #
        self._start_browser(args.show_browser)
        if not self._login():
            sys.exit()

        if args.status:
            # Look up current mentee status for each mentor
            #
            start_time = time.time()
            self._make_dir(args.status)
            self._get_current_mentee_status(args.status)
            print('Status completed in {0:.0f} seconds. Written to {1}\n'.format(time.time() - start_time, args.status))

        if args.input:
            # Load the "daily report" xls
            #
            start_time = time.time()
            kitten_report_reader = KittenReportReader(self._dog_mode)
            if not kitten_report_reader.open_xls(args.input):
                sys.exit()

            # Process the daily report
            #
            animal_numbers = kitten_report_reader.get_animal_numbers()
            print('Found {} animal numbers: {}'.format(len(animal_numbers), ', '.join([str(a) for a in animal_numbers])))

            # Query details for each animal (current foster parent, foster status, etc)
            #
            foster_parents, animal_details, filtered_animals = self._get_animal_details(animal_numbers)

            for p in foster_parents:
                print('Animals for foster parent {} = {}'.format(p, foster_parents[p]))

            # Query foster parent details (person number -> name, contact details, etc)
            #
            persons_data = {}
            for person in foster_parents:
                persons_data[person] = self._get_person_data(person)

            # Output the combined results to csv
            #
            self._make_dir(args.output)
            kitten_report_reader.output_results(persons_data, foster_parents, animal_details, filtered_animals, args.output)

            print('\n{0} foster report completed in {1:.0f} seconds. Written to {2}'.format(
                'Feline' if not self._dog_mode else 'Canine',
                time.time() - start_time,
                args.output))

        self._exit_browser()

    def _load_config_file(self):
        ''' A config.yaml configuration file is expected to be in the same directory as this script
        '''
        try:
            config_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'config.yaml')
            config = yaml.load(open(config_file, 'r'), Loader=yaml.SafeLoader)
            self._username = config['username']
            self._password = config['password']
            self._dog_mode = config['dog_mode'] if 'dog_mode' in config else False

            self._google_spreadsheet_key = config['google_spreadsheet_key'] if 'google_spreadsheet_key' in config else None
            self._google_client_secret = config['google_client_secret'] if 'google_client_secret' in config else None

            self._box_user_id = config['box_user_id'] if 'box_user_id' in config else None
            self._box_file_id = config['box_file_id'] if 'box_file_id' in config else None
            self._box_jwt = config['box_jwt'] if 'box_jwt' in config else None

            if not (self._google_spreadsheet_key and self._google_client_secret) and not (self._box_user_id and self._box_file_id and self._box_jwt):
                print_err('ERROR: Incomplete mentor spreadsheet configuration: {}'.format(config_file))
                return False

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

    def _read_additional_config_yaml(self, additional_config_yaml):
        ''' The mentors spreadsheet contains additional configuration data. This makes it easier to manage dynamic
            configuration data vs rollout of config.yaml updates.
        '''
        try:
            config = yaml.load(additional_config_yaml, Loader=yaml.SafeLoader)
            self._login_url = config['login_url']
            self._search_url = config['search_url']
            self._animal_url = config['animal_url']
            self._medical_details_url = config['medical_details_url']
            self._list_animals_url = config['list_animals_url']
            self._responsible_for_url = config['responsible_for_url']
            self._do_not_assign_mentor = config['do_not_assign_mentor'] if 'do_not_assign_mentor' in config else []
            self._mentors = config['mentors'] if 'mentors' in config else []

        except AttributeError as err:
            print_err('ERROR: Unable to read additional config: {}'.format(err))
            return False

        except yaml.YAMLError as err:
            print_err('ERROR: Unable to read additional config: {}'.format(err))
            return False

        except TypeError as err:
            print_err('ERROR: Invalid yaml in additional config: {}'.format(err))
            return False

        except KeyError as err:
            print_err('ERROR: Missing value in additional config: {}'.format(err))
            return False

        return True

    def _start_browser(self, show_browser):
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
        elif sys.platform.startswith('win32'):
            chromedriver_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bin/win32/chromedriver')
            chrome_options.add_experimental_option('excludeSwitches', ['enable-logging']) # chromedriver complains a lot on Windows
        else:
            print_err('Sorry friends, I haven\'t included chromedriver for your platform ({}). Exiting now.'.format(sys.platform))
            sys.exit(0)

        self._driver = webdriver.Chrome(chromedriver_path, options = chrome_options)
        self._driver.set_page_load_timeout(60)

    def _exit_browser(self):
        ''' Close and exit the browser
        '''
        self._driver.close()
        self._driver.quit()

    def _login(self):
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

    def _get_animal_details(self, animal_numbers):
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

                # Remove empty lines and double quotes
                #
                special_msg = os.linesep.join([s for s in special_msg.splitlines() if s])
                special_msg = special_msg.replace('"', '\'')

            animal_details[a] = {}
            animal_details[a]['message'] = special_msg
            status = self._get_selection_by_id('status')
            sub_status = self._get_selection_by_id('subStatus')
            animal_details[a]['status'] = '{}{}{}'.format(status, ' - ' if sub_status else '', sub_status)
            animal_details[a]['name'] = self._get_attr_by_id('animalname').strip()

            try: 
                animal_details[a]['status_date'] = datetime.strptime(self._get_attr_by_id('statusdate'), '%m/%d/%Y').strftime('%-d-%b-%Y')
            except ValueError:
                animal_details[a]['status_date'] = 'Unknown'

            # If this animal is currently in foster, get the responsible person (foster parent)
            # Dog mode: include all status other than 'adopted'
            #
            status = status.lower()
            if (self._dog_mode and 'adopted' not in status) or ('in foster' in status and 'unassisted death' not in status):
                try:
                    p = int(self._get_attr_by_xpath('href', '//*[@id="Table17"]/tbody/tr[1]/td[2]/a').split('personid=')[1])
                    foster_parents.setdefault(p, []).append(a)
                except:
                    print_err('Failed to find foster parent for animal {}, please check report'.format(a))
            else:
                filtered_animals.add(a)

            animal_details[a]['sn'] = self._get_spay_neuter_status(a)
            print('{}'.format(animal_details[a]['status']))

        return foster_parents, animal_details, filtered_animals

    def _get_spay_neuter_status(self, animal_number):
        # Load spay/neuter status from the medical details page
        #
        try:
            self._driver.get(self._medical_details_url.format(animal_number))
            return utf8(self._get_attr_by_xpath('innerText', '/html/body/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[4]'))
        except:
            print_err('Failed to read spay/neuter status for animal {}'.format(animal_number))
            return 'Unknown'

    def _get_person_data(self, person_number):
        ''' Load the given person number, return details and contact information
        '''
        print('Looking up person {}... '.format(person_number), end='')
        sys.stdout.flush()

        self._driver.get(self._search_url)
        self._driver.find_element_by_id('userid').send_keys(str(person_number))
        self._driver.find_element_by_id('userid').send_keys(webdriver.common.keys.Keys.RETURN)

        first_name     = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtFirstName').strip()
        last_name      = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtLastName').strip()
        preferred_name = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtPreferredName').strip()
        home_phone     = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_homePhone_txtPhone3').strip()
        cell_phone     = self._get_attr_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_mobilePhone_txtPhone3').strip()
        email1         = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[1]/td[1]').strip()
        email2         = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[2]/td[1]').strip()
        email3         = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[3]/td[1]').strip()
        email4         = self._get_attr_by_xpath('innerText', '//*[@id="emailTable"]/tbody/tr[4]/td[1]').strip()

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
        matching_sheets = self.mentor_sheet_reader.find_matching_mentors(match_strings)
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
                fostered_tr_active = False
                agency_outgoing_tr_active = False

                for row in rows:
                    cols = row.find_elements(By.TAG_NAME, 'td')
                    num_cols = len(cols)

                    if num_cols == 10:
                        animal_type = cols[5].text.lower()
                        animal_status = cols[2].text.lower()
                        if fostered_tr_active or agency_outgoing_tr_active:
                            target_types = ['cat', 'kitten'] if not self._dog_mode else ['dog', 'puppy']
                            if any(s in animal_type for s in target_types):
                                if 'in foster' not in animal_status or animal_status == 'unassisted death - in foster':
                                    previous_feline_foster_count += 1

                                if 'euthanized' in animal_status:
                                    euthanized_count += 1
                                elif 'unassisted death' in animal_status:
                                    unassisted_death_count += 1

                    elif num_cols == 1:
                        fostered_tr_active = cols[0].text.lower() == 'fostered'
                        agency_outgoing_tr_active = cols[0].text.lower() == 'agency outgoing'

                    else:
                        fostered_tr_active = False
                        agency_outgoing_tr_active = False

            except NoSuchElementException:
                break

            page_number = page_number + 1
        return previous_feline_foster_count, euthanized_count, unassisted_death_count

    def _print_and_write(self, file, s):
        print(s)
        file.write('{}\r\n'.format(s))

    def _get_current_mentee_status(self, outfile):
        ''' Retrieve animals in foster for all current mentees
        '''
        with open(outfile, 'w') as f:
            current_mentees = self.mentor_sheet_reader.get_current_mentees()
            for current in current_mentees:
                self._print_and_write(f, '-------------------------------------------')
                self._print_and_write(f, current['mentor'])
                if len(current['mentees']):
                    for mentee in current['mentees']:
                        current_animals = self._current_animals_fostered(mentee['name'], mentee['pid'])
                        self._print_and_write(f, '    {} ({}) - {} animals'.format(mentee['name'].replace('\n', ' '), mentee['pid'], len(current_animals)))
                        for a in current_animals:
                            self._print_and_write(f, ('        {} (S/N {})'.format(a, self._get_spay_neuter_status(a))))
                else:
                    self._print_and_write(f, '    ** No current mentees **')

                self._print_and_write(f, '')

    def _current_animals_fostered(self, person_name, person_number):
        current_animals = []
        self._driver.get(self._responsible_for_url.format(person_number))
        try:
            table = self._driver.find_element_by_id('Table1')
            for row in table.find_elements(By.TAG_NAME, 'tr'):
                cols = row.find_elements(By.TAG_NAME, 'td')
                if len(cols) == 12:
                    animal_status = cols[2].text.lower()
                    if 'in foster' in animal_status and animal_status != 'unassisted death - in foster':
                        animal_number = int(cols[3].text)
                        if animal_number not in current_animals: # ignore duplicates
                            current_animals.append(animal_number)
        except NoSuchElementException:
            pass

        return current_animals

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

    def _make_dir(self, fullpath):
        dirname = os.path.dirname(fullpath)
        if dirname and not os.path.exists(dirname):
            os.makedirs(dirname)

if __name__ == "__main__":
    KittenScraper().run()
