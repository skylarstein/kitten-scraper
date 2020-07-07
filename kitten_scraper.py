from __init__ import __version__
import os
import re
import math
import sys
import time
import yaml
from argparse import ArgumentParser
from datetime import date, datetime
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
        start_time = time.time()

        arg_parser = ArgumentParser()
        arg_parser.add_argument('-i', '--input', help = 'specify the daily kitten report (xls), or optionally a comma-separated list of animal numbers', required = False)
        arg_parser.add_argument('-s', '--mentee_status', help = 'retrieve current mentee status [verbose,autoupdate,export]', required = False, nargs='?', default='', const='yes')
        arg_parser.add_argument('-c', '--config', help = 'specify a config file', required = True)
        arg_parser.add_argument('-b', '--show_browser', help = 'show the browser window (generally for debugging)', required = False, action = 'store_true')
        args = arg_parser.parse_args()

        if not args.input and not args.mentee_status:
            arg_parser.print_help()
            sys.exit(0)

        # Load config.yaml
        #
        if not self._load_config_file(args.config):
            sys.exit()

        # Load the Foster Mentors spreadsheet
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

        # Load additional config data from the mentors spreadsheet. This minimizes the need to deploy updates to the
        # local config.yaml file.
        #
        if not self._read_additional_config_yaml(self._additional_config_yaml):
            sys.exit()

        # Start the browser, log in
        #
        self._start_browser(args.show_browser)
        if not self._login():
            sys.exit()

        current_mentee_status = self._get_current_mentee_status(args.mentee_status) if args.mentee_status else None

        if args.input:
            # Load animal numbers. Note that args.input will either be a path to the "daily report" xls, or may
            # optionally be a comma-separated list of animal numbers.
            #
            if re.fullmatch(r'(\s?\d+\s?)(\s?,\s?\d+\s?)*$', args.input):
                animal_numbers = [s.strip() for s in args.input.split(',')]
            else:
                animal_numbers = KittenReportReader().read_animal_numbers_from_xls(args.input)
    
            if not animal_numbers:
                sys.exit()

            print('Found {} animal{}: {}'.format(
                len(animal_numbers),
                's' if len(animal_numbers) != 1 else '',
                ', '.join([str(a) for a in animal_numbers])))

            # Query details for each animal (current foster parent, foster status, breed, color, gender, age, etc.)
            #
            animal_data, foster_parents, animals_not_in_foster = self._get_animal_data(animal_numbers)

            for p in foster_parents:
                print('Animals for foster parent {} = {}'.format(p, foster_parents[p]))

            # Query details for each foster parent (name, contact details, etc.)
            #
            persons_data = {}
            for person in foster_parents:
                persons_data[person] = self._get_person_data(person)

            # Export mentee status to file (if '--mentee_status export')
            #
            if current_mentee_status:
                status_file = os.path.join(default_dir(), '{}_foster_mentor_status_{}.txt'.format(self.BASE_ANIMAL_TYPE, date.today().strftime('%Y.%m.%d')))
                with open(status_file, 'w') as f:
                    for current in current_mentee_status:
                        self._print_and_write(f, '--------------------------------------------------')
                        self._print_and_write(f, current['mentor'])
                        if len(current['mentees']):
                            for mentee in current['mentees']:
                                self._print_and_write(f, '    {} ({}) - {} animals'.format(mentee['name'],  mentee['pid'], len(mentee['current_animals'])))
                                for a, data in mentee['current_animals'].items():
                                    if 'sn' in data: # queried/populated if '--mentee_status verbose'
                                        self._print_and_write(f, '        {} (S/N {})'.format(a, data['sn']))
                                    else:
                                        self._print_and_write(f, '        {}'.format(a))
                        else:
                            self._print_and_write(f, '    ** No current mentees **')

                        self._print_and_write(f, '')

            # Save report to file
            #
            output_csv = os.path.join(default_dir(), '{}_foster_mentor_report_{}.csv'.format(self.BASE_ANIMAL_TYPE, date.today().strftime('%Y.%m.%d')))
            self._make_dir(output_csv)
            self._output_results(animal_data,
                                 foster_parents,
                                 persons_data,
                                 animals_not_in_foster,
                                 current_mentee_status,
                                 output_csv)

        print('KittenScraper completed in {0:.0f} seconds'.format(time.time() - start_time))
        self._exit_browser()

    def _start_browser(self, show_browser):
        ''' Instantiate the browser, configure options as needed
        '''
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/605.1.15 '
                                    '(KHTML, like Gecko) Version/12.0.3 Safari/605.1.15')
        if not show_browser:
            chrome_options.add_argument('--headless')

        # TODO: Consider adding chromedriver-binary or chromedriver_installer to requirements.txt and
        # removing these local copies.
        #
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
            self._driver.find_element_by_id('txt_username').send_keys(self._username)
            self._driver.find_element_by_id('txt_password').send_keys(self._password)
            self._driver.find_element_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_btn_login').click()
            self._driver.find_element_by_id('Continue').click()

        except NoSuchElementException:
            print_err('ERROR: Unable to login. Please check your username/password.')
            return False

        return True

    def _load_config_file(self, config_file_yaml):
        ''' A config.yaml configuration file is expected to be in the same directory as this script
        '''
        try:
            config_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), config_file_yaml)
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

            if self._dog_mode:
                print_warn('** Dog Mode is Active **')

            self.YOUNG_ANIMAL_TYPE = 'kitten' if not self._dog_mode else 'puppy'
            self.ADULT_ANIMAL_TYPE = 'cat' if not self._dog_mode else 'dog'
            self.BASE_ANIMAL_TYPE = 'feline' if not self._dog_mode else 'canine'

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

    def _get_animal_data(self, animal_numbers):
        ''' Load additional animal data for each animal number
        '''
        animal_data = {}
        foster_parents = {}
        animals_not_in_foster = set()
        for a in animal_numbers:
            print('Looking up animal {}... '.format(a), end='', flush=True)
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

            animal_data[a] = {}
            animal_data[a]['message'] = special_msg
            status = self._get_selection_by_id('status')

            if not status:
                # Status text is usually found within a <select> element, but is sometimes found as innerText within the
                # <td> that looks something like this: "Adopted - Awaiting Pickup\nChange Status"
                try:
                    status = self._get_property_by_xpath('innerText', '//*[@id="Table17"]/tbody/tr[5]/td[2]').split('\n')[0]
                except:
                    status = ''

            sub_status = self._get_selection_by_id('subStatus')
            animal_data[a]['status'] = '{}{}{}'.format(status, ' - ' if sub_status else '', sub_status)
            animal_data[a]['name'] = self._get_attr_by_id('animalname').strip()
            animal_data[a]['breed'] = self._get_attr_by_id('primaryBreed').strip()
            animal_data[a]['primary_color'] = self._get_selection_by_id('primaryColour')
            animal_data[a]['secondary_color'] = self._get_selection_by_id('secondaryColour')
            animal_data[a]['gender'] = self._get_selection_by_id('sex')

            try:
                age = datetime.now() - datetime.strptime(self._get_attr_by_id('dob'), '%m/%d/%Y')
                animal_data[a]['age'], animal_data[a]['type'] = self._stringify_age_and_type(age)
            except:
                animal_data[a]['age'] = 'Unknown Age'
                animal_data[a]['type'] = self.BASE_ANIMAL_TYPE

            try:
                animal_data[a]['status_date'] = datetime.strptime(self._get_attr_by_id('statusdate'), '%m/%d/%Y').strftime('%-d-%b-%Y')
            except ValueError:
                animal_data[a]['status_date'] = 'Unknown'

            # If this animal is currently in foster, get the responsible person (foster parent).
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
                animals_not_in_foster.add(a)

            animal_data[a]['sn'] = self._get_spay_neuter_status(a) # Do this last, this will load a new web page!

            # Create some helpful/default string representations
            #
            set_default = lambda str, default : default if str.strip() in [None, ''] else str

            animal_data[a]['sn'] = set_default(animal_data[a]['sn'], 'Unknown')
            animal_data[a]['status'] = set_default(animal_data[a]['status'], 'Status Unknown')
            animal_data[a]['name'] = set_default(animal_data[a]['name'], 'Unnamed')
            animal_data[a]['breed'] = set_default(animal_data[a]['breed'], 'Unknown Breed')
            animal_data[a]['gender'] = set_default(animal_data[a]['gender'], 'Unknown Gender')
            animal_data[a]['color'] = set_default(animal_data[a]['primary_color'], 'Unknown Color')

            if animal_data[a]['secondary_color'].strip() not in [None, '', 'None']:
                animal_data[a]['color'] += '/{}'.format(animal_data[a]['secondary_color'])

            breed_abbreviations = {
                # 'Breed' is a free-form text field and not everyone enters data exactly the same. I'll compare against
                # lowercase/no-whitespace for slightly better odds of matches.
                #
                'domesticshorthair'  : 'DSH',
                'domesticmediumhair' : 'DMH',
                'domesticlonghair'   : 'DLH' }

            abbreviation = breed_abbreviations.get(animal_data[a]['breed'].replace(' ', '').lower())
            if abbreviation:
                animal_data[a]['breed'] = abbreviation

            if animal_data[a]['gender'].lower() == 'male':
                animal_data[a]['gender_short'] = 'M'
            elif animal_data[a]['gender'].lower() == 'female':
                animal_data[a]['gender_short'] = 'F'
            else:
                animal_data[a]['gender_short'] = animal_data[a]['gender']

            print('{}'.format(animal_data[a]['status']))

        return animal_data, foster_parents, animals_not_in_foster

    def _stringify_age_and_type(self, age):
        ''' To keep things brief and easy to read, I'll floor()/round() months and weeks. This is close enough for
            informational purposes.
        '''
        try:
            if age.days >= 365:
                years = math.floor(age.days / 365)
                months = round((age.days % 365) / 30)
                age_string = '{:.0f} year{}, {:.0f} month{}'.format(years,
                                                                    's' if years > 1 else '',
                                                                    months,
                                                                    's' if months != 1 else '')
                animal_type = self.ADULT_ANIMAL_TYPE

            elif age.days >= 90:
                months = math.floor(age.days / 30)
                age_string = '{} month{}'.format(months, 's' if months > 1 else '')
                animal_type = self.YOUNG_ANIMAL_TYPE if age.days <= (365 / 2) else ADULT_ANIMAL_TYPE

            else:
                weeks = math.floor(age.days / 7)
                age_string = '{} week{}'.format(weeks, 's' if weeks > 1 else '')
                animal_type = self.YOUNG_ANIMAL_TYPE
        except:
            animal_type = self.BASE_ANIMAL_TYPE
            age_string = 'Unknown Age'

        return age_string, animal_type

    def _get_spay_neuter_status(self, animal_number):
        ''' Load spay/neuter status from the medical details page
        '''
        try:
            self._driver.get(self._medical_details_url.format(animal_number))
            return utf8(self._get_attr_by_xpath('innerText', '/html/body/center/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[4]'))
        except:
            print_err('Failed to read spay/neuter status for animal {}'.format(animal_number))
            return 'Unknown'

    def _get_person_data(self, person_number):
        ''' Load the given person number, return details and contact information
        '''
        print('Looking up person {}... '.format(person_number), end='', flush=True)
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

        match_strings = emails.union([full_name, str(person_number)])
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
        ''' Determine the total number of animals this person has previously fostered. This is a useful metric to gauge
            experience level, but there are some difficulties interpreting the data without getting unnecessarily crazy
            in here. Consider these numbers "a decent guess". 

            Load the list of all animals this person has been responsible for, page by page until we have no more pages.
        '''
        page_number = 1
        previous_foster_count = 0
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
                                    previous_foster_count += 1

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
        return previous_foster_count, euthanized_count, unassisted_death_count

    def _print_and_write(self, file, s):
        print(s)
        file.write('{}\r\n'.format(s))

    def _get_current_mentee_status(self, arg_mentee_status):
        ''' Get current mentees and mentee status for each mentor
        '''
        verbose_status = 'verbose' in arg_mentee_status # print status of each animail assigned to a mentor
        autoupdate_completed_mentees = 'autoupdate' in arg_mentee_status # mark 'completed' mentors in the spreadsheet
        print_success('Looking up mentee status (verbose_status = {}, autoupdate_completed_mentees = {})...'.format(
            verbose_status,
            autoupdate_completed_mentees))

        completed_mentees = {}
        current_mentees = self.mentor_sheet_reader.get_current_mentees()
        for current in current_mentees:
            current['active_count'] = 0
            print('Checking mentee status for {}... '.format(current['mentor']), end='', flush=True)

            if len(current['mentees']):
                for mentee in current['mentees']:
                    current_animal_ids = self._current_animals_fostered(mentee['name'], mentee['pid'])
                    mentee['current_animals'] = {}

                    for current_animal_id in current_animal_ids:
                        if verbose_status:
                            mentee['current_animals'][current_animal_id] = { 'sn' : self._get_spay_neuter_status(current_animal_id) }
                        else:
                            mentee['current_animals'][current_animal_id] = { }

                    if len(current_animal_ids):
                        current['active_count'] = current['active_count'] + 1
                    else:
                        completed_mentees.setdefault(current['mentor'], []).append(mentee['pid'])

            days_ago = (datetime.now() - current['most_recent']).days if current['most_recent'] else 'N/A'
            print('active mentees = {}, last assigned days ago = {}'.format(current['active_count'], days_ago))

        if autoupdate_completed_mentees:
            print_success('Auto-updating completed mentees in the mentor spreadsheet...')
            for mentor in completed_mentees.keys():
                self.mentor_sheet_reader.set_completed_mentees(mentor, completed_mentees[mentor])

        return current_mentees

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

    def _output_results(self, animal_data, foster_parents, persons_data, animals_not_in_foster, current_mentee_status, csv_filename):
        ''' Output all of our new super amazing results to a csv file
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
        csv_rows[-1].append('Date {}s Received'.format(self.BASE_ANIMAL_TYPE.capitalize()))
        if self._dog_mode:
            csv_rows[-1].append('"Name, Breed, Color"')
        csv_rows[-1].append('Animal Details')
        csv_rows[-1].append('Special Animal Message')

        # Build a row for each foster parent
        #
        for person_number in sorted(foster_parents, key=lambda p: persons_data[p]['notes']):
            person_data = persons_data[person_number]
            name = person_data['full_name']
            report_notes = person_data['notes']
            loss_rate = round(person_data['loss_rate'])
            animals_with_this_person = foster_parents[person_number]
            animal_details, animal_details_brief = self._get_animal_details_string(animals_with_this_person, animal_data)

            prev_animals_fostered = person_data['prev_animals_fostered']
            foster_experience = 'NEW' if not prev_animals_fostered else '{}'.format(prev_animals_fostered)

            special_message = ''
            for a in animals_with_this_person:
                msg = animal_data[a]['message']
                if msg:
                    special_message += '{}{}: {}'.format('\r\r' if special_message else '', a, msg)

            cell_number = person_data['cell_phone']
            home_number = person_data['home_phone']
            phone = ''
            if len(cell_number) >= 10: # ignore incomplete phone numbers
                phone = '(C) {}'.format(cell_number)
            if len(home_number) >= 10: # ignore incomplete phone numbers
                phone += '{}(H) {}'.format('\r' if phone else '', home_number)

            email = ''
            for e in person_data['emails']:
                email += '{}{}'.format('\r' if email else '', e)

            # I will assume all animals in this group went into foster on the same date. This should usually be true
            # since this is designed to processed with a "daily report".
            #
            date_received = animal_data[animals_with_this_person[0]]['status_date']

            # Explicitly wrap numbers/datestr with ="{}" to avoid Excel auto-formatting issues
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
            if self._dog_mode:
                csv_rows[-1].append('"{}"'.format(animal_details_brief))
            csv_rows[-1].append('"{}"'.format(animal_details))
            csv_rows[-1].append('"{}"'.format(special_message))

            print('{} (Experience: {}, Loss Rate: {}%) {}{}{}'.format(name,
                                                                      foster_experience, 
                                                                      loss_rate,
                                                                      ConsoleFormat.GREEN,
                                                                      report_notes.replace('\r', ', '),
                                                                      ConsoleFormat.END))
        with open(csv_filename, 'w') as outfile:
            for row in csv_rows:
                outfile.write(','.join(row))
                outfile.write('\n')

            if not len(foster_parents):
                outfile.write('*** None of the animals in this report are currently in foster\n')
                print_warn('None of the animals in this report are currently in foster. Nothing to do!')

            if len(animals_not_in_foster):
                outfile.write('\n\n\n*** Animals not in foster\n')
                print_warn('\nAnimals not in foster')
                for a in animals_not_in_foster:
                    outfile.write('{} - {}\n'.format(a, animal_data[a]['status']))
                    print('{} - {}'.format(a, animal_data[a]['status']))

            if current_mentee_status:
                outfile.write('\n\nMentor,Active Mentees,Last Assigned (days ago)\n')
                for current in current_mentee_status:
                    days_ago = (datetime.now() - current['most_recent']).days if current['most_recent'] else 'N/A'
                    outfile.write('{},{},{}\n'.format(current['mentor'], current['active_count'], days_ago))

    def _get_animal_details_string(self, foster_animals, animal_data):
        ''' Group animals by type, list useful details for each animal
        '''
        animals_by_type = {}
        for a in foster_animals:
            animals_by_type.setdefault(animal_data[a]['type'], []).append(a)

        animal_details = ''
        for animal_type in animals_by_type.keys():
            animals = animals_by_type[animal_type]

            line = '{} {}{} @ {}'.format(len(animals),
                                         animal_type, 
                                         's' if len(animals) > 1 else '',
                                         animal_data[animals[0]]['age'])
            for a in sorted(animals):
                line += '\r{} ({}), S/N {}'.format(a,
                                                   animal_data[a]['gender_short'],
                                                   animal_data[a]['sn'])
                if not self._dog_mode:
                    line += ', {}, {}, {}'.format(animal_data[a]['name'],
                                                  animal_data[a]['breed'],
                                                  animal_data[a]['color'])

            animal_details += '{}{}'.format('\r' if animal_details else '', line)

        animal_details_brief = ''
        for a in sorted(foster_animals):
            animal_details_brief += '{}{} - {}, {}, {}'.format('\r' if animal_details_brief else '',
                                                               a,
                                                               animal_data[a]['name'],
                                                               animal_data[a]['breed'],
                                                               animal_data[a]['color'])
        return animal_details, animal_details_brief

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

    def _get_property_by_xpath(self, property, element_xpath):
        try:
            return self._driver.find_element_by_xpath(element_xpath).get_property(property)
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
