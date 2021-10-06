from argparse import ArgumentParser
from datetime import date, datetime
import os
import re
import math
import sys
import time
import yaml
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from __init__ import __version__
from box_sheet_reader import BoxSheetReader
from google_sheet_reader import GoogleSheetReader
from kitten_report_reader import KittenReportReader
from kitten_utils import Log, Utils

class KittenScraper(object):
    def __init__(self):
        self.mentor_sheet_reader = None
        self._additional_config_yaml = None

    def run(self):
        print('Welcome to KittenScraper {}'.format(__version__))
        start_time = time.time()

        arg_parser = ArgumentParser()
        arg_parser.add_argument('-i', '--input', help = 'specify the daily foster report (xls), or optionally a comma-separated list of animal numbers', required = False)
        arg_parser.add_argument('-s', '--mentee_status', help = 'retrieve current mentee status [autoupdate,export]', required = False, nargs='?', default='', const='yes')
        arg_parser.add_argument('-c', '--config', help = 'specify a config file (optional, defaults to \'config.yaml\')', required = False, default='config.yaml')
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
            Log.error('ERROR: Incorrect mentor spreadsheet configuration, please check config.yaml')
            sys.exit()

        if self._additional_config_yaml is None:
            Log.error('ERROR: configuration YAML from mentors spreadsheet not found, cannot continue')
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

        if current_mentee_status:
            status_file = None
            export_status = 'export' in args.mentee_status

            if export_status:
                status_file_path = os.path.join(Utils.default_dir(), '{}_foster_mentor_status_{}.txt'.format(self.BASE_ANIMAL_TYPE, date.today().strftime('%Y.%m.%d')))
                status_file = open(status_file_path, 'w')
                Log.success(f'Exporting mentee status to file: {status_file_path}')

            for current in current_mentee_status:
                self._print_and_write(status_file, '--------------------------------------------------')
                self._print_and_write(status_file, current['mentor'])
                if current['mentees']:
                    for mentee in current['mentees']:
                        self._print_and_write(status_file, '    {} ({}) - {} animals'.format(mentee['name'],  mentee['pid'], len(mentee['current_animals'])))
                        for a_number, data in mentee['current_animals'].items():
                            self._print_and_write(status_file, '        {} (S/N {}, Bio {}, Photo {})'.format(a_number, data['sn'], data['bio'], data['photo']))
                else:
                    self._print_and_write(status_file, '    ** No current mentees **')

                self._print_and_write(status_file, '')

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

            for p_number in foster_parents:
                print('Animals for foster parent {} = {}'.format(p_number, foster_parents[p_number]))

            # Query details for each foster parent (name, contact details, etc.)
            #
            persons_data = {}
            for person in foster_parents:
                persons_data[person] = self._get_person_data(person)

            # Save report to file
            #
            output_csv = os.path.join(Utils.default_dir(), '{}_foster_mentor_report_{}.csv'.format(self.BASE_ANIMAL_TYPE, date.today().strftime('%Y.%m.%d')))
            Utils.make_dir(output_csv)
            self._output_results(animal_data,
                                 foster_parents,
                                 persons_data,
                                 animals_not_in_foster,
                                 current_mentee_status,
                                 output_csv)

            # Optional: automatically forward this report via email
            #
            if 'generate_email' in self.config:
                from outlook_email import compose_outlook_email
                subject = self._get_from_dict(self.config['generate_email'], 'subject')
                recipient_name = self._get_from_dict(self.config['generate_email'], 'recipient_name')
                recipient_email = self._get_from_dict(self.config['generate_email'], 'recipient_email')
                message = self._get_from_dict(self.config['generate_email'], 'message')

                if None not in [subject, recipient_name, recipient_email, message]:
                    compose_outlook_email(subject=subject,
                                          recipient_name=recipient_name,
                                          recipient_email=recipient_email,
                                          body=message,
                                          attachment=output_csv)
                    Log.debug('Composed email to {} <{}>'.format(recipient_name, recipient_email))

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

        # Consider adding chromedriver-binary or chromedriver_installer to requirements.txt and
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
            Log.error('Sorry friends, I haven\'t included chromedriver for your platform ({}). Exiting now.'.format(sys.platform))
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
        Log.success('Logging in...')

        try:
            self._driver.set_page_load_timeout(20)
            self._driver.get(self._login_url)

        except TimeoutException:
            Log.error('ERROR: Unable to load the login page. Please check your connection.')
            return False
        except NoSuchElementException:
            Log.error('ERROR: Unable to load the login page. Please check your connection.')
            return False

        try:
            self._driver.find_element_by_id('txt_username').send_keys(self._username)
            self._driver.find_element_by_id('txt_password').send_keys(self._password)
            self._driver.find_element_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_btn_login').click()
            self._driver.find_element_by_id('Continue').click()

        except NoSuchElementException:
            Log.error('ERROR: Unable to login. Please check your username/password.')
            return False

        return True

    def _load_config_file(self, config_file_yaml):
        ''' A config.yaml configuration file is expected to be in the same directory as this script
        '''
        try:
            config_file = os.path.join(os.path.dirname(os.path.realpath(__file__)), config_file_yaml)
            self.config = yaml.load(open(config_file, 'r'), Loader=yaml.SafeLoader)
            self._username = self.config['username']
            self._password = self.config['password']
            self._dog_mode = self.config['dog_mode'] if 'dog_mode' in self.config else False

            self._google_spreadsheet_key = self.config['google_spreadsheet_key'] if 'google_spreadsheet_key' in self.config else None
            self._google_client_secret = self.config['google_client_secret'] if 'google_client_secret' in self.config else None

            self._box_user_id = self.config['box_user_id'] if 'box_user_id' in self.config else None
            self._box_file_id = self.config['box_file_id'] if 'box_file_id' in self.config else None
            self._box_jwt = self.config['box_jwt'] if 'box_jwt' in self.config else None

            if not (self._google_spreadsheet_key and self._google_client_secret) and not (self._box_user_id and self._box_file_id and self._box_jwt):
                Log.error('ERROR: Incomplete mentor spreadsheet configuration: {}'.format(config_file))
                return False

            if self._dog_mode:
                Log.warn('** Dog Mode is Active **')

            self.YOUNG_ANIMAL_TYPE = 'kitten' if not self._dog_mode else 'puppy'
            self.ADULT_ANIMAL_TYPE = 'cat' if not self._dog_mode else 'dog'
            self.BASE_ANIMAL_TYPE = 'feline' if not self._dog_mode else 'canine'

        except yaml.YAMLError as err:
            Log.error('ERROR: Unable to parse configuration file: {}, {}'.format(config_file, err))
            return False

        except IOError as err:
            Log.error('ERROR: Unable to read configuration file: {}, {}'.format(config_file, err))
            return False

        except KeyError as err:
            Log.error('ERROR: Missing value in configuration file: {}, {}'.format(config_file, err))
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
            self._list_all_animals_url = config['list_animals_url']
            self._responsible_for_url = config['responsible_for_url']
            self._responsible_for_paged_url = config['responsible_for_paged_url']
            self._adoption_summary_url = config['adoption_summary_url']
            self._do_not_assign_mentor = config['do_not_assign_mentor'] if 'do_not_assign_mentor' in config else []
            self._mentors = config['mentors'] if 'mentors' in config else []

        except AttributeError as err:
            Log.error('ERROR: Unable to read additional config: {}'.format(err))
            return False

        except yaml.YAMLError as err:
            Log.error('ERROR: Unable to read additional config: {}'.format(err))
            return False

        except TypeError as err:
            Log.error('ERROR: Invalid yaml in additional config: {}'.format(err))
            return False

        except KeyError as err:
            Log.error('ERROR: Missing value in additional config: {}'.format(err))
            return False

        return True

    def _get_animal_data(self, animal_numbers, silent = False):
        ''' Load additional animal data for each animal number
        '''
        animal_data = {}
        foster_parents = {}
        animals_not_in_foster = set()
        for a_number in animal_numbers:
            if not silent:
                print('Looking up animal {}... '.format(a_number), end='', flush=True)
            sys.stdout.flush()
            self._driver.get(self._animal_url.format(a_number))
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
            except Exception:
                raise Exception('Timeout while waiting for content on search page!') from Exception

            # Get Special Message text (if it exists)
            #
            special_msg = Utils.utf8(self._get_text_by_id('specialMessagesDialog'))

            if special_msg:
                # Remove text we don't care about
                #
                special_msg = re.sub(r'(?i)This is a special message. If you would like to delete it then clear the Special Message box in the General Details section of this page.', '', special_msg).strip()

                # Remove empty lines and double quotes
                #
                special_msg = os.linesep.join([s for s in special_msg.splitlines() if s])
                special_msg = special_msg.replace('"', '\'')

            animal_data[a_number] = {}
            animal_data[a_number]['message'] = special_msg
            status = self._get_selection_by_id('status')

            if not status:
                # Status text is usually found within a <select> element, but is sometimes found as innerText within the
                # <td> that looks something like this: "Adopted - Awaiting Pickup\nChange Status"
                try:
                    status = self._get_property_by_xpath('innerText', '//*[@id="Table17"]/tbody/tr[5]/td[2]').split('\n')[0]
                except Exception:
                    status = ''

            sub_status = self._get_selection_by_id('subStatus')
            animal_data[a_number]['status'] = '{}{}{}'.format(status, ' - ' if sub_status else '', sub_status)
            animal_data[a_number]['name'] = self._get_attr_by_id('animalname').strip()
            animal_data[a_number]['breed'] = self._get_attr_by_id('primaryBreed').strip()
            animal_data[a_number]['primary_color'] = self._get_selection_by_id('primaryColour')
            animal_data[a_number]['secondary_color'] = self._get_selection_by_id('secondaryColour')
            animal_data[a_number]['gender'] = self._get_selection_by_id('sex')
            animal_data[a_number]['photo'] = 'No' if 'NoImage.png' in self._get_property_by_xpath('src', '//*[@id="animal-default-photo"]') else 'Yes'

            try:
                age = datetime.now() - datetime.strptime(self._get_attr_by_id('dob'), '%m/%d/%Y')
                animal_data[a_number]['age'], animal_data[a_number]['type'] = self._stringify_age_and_type(age)
            except Exception:
                animal_data[a_number]['age'] = 'Unknown Age'
                animal_data[a_number]['type'] = self.BASE_ANIMAL_TYPE

            try:
                animal_data[a_number]['status_date'] = datetime.strptime(self._get_attr_by_id('statusdate'), '%m/%d/%Y').strftime('%-d-%b-%Y')
            except ValueError:
                animal_data[a_number]['status_date'] = 'Unknown'

            # If this animal is currently in foster, get the responsible person (foster parent).
            #
            status = status.lower()
            if ('in foster' in status and 'unassisted death' not in status):
                try:
                    p_number = int(self._get_attr_by_xpath('href', '//*[@id="Table17"]/tbody/tr[1]/td[2]/a').split('personid=')[1])
                    foster_parents.setdefault(p_number, []).append(a_number)
                except Exception:
                    Log.error('Failed to find foster parent for animal {}, please check report'.format(a_number))
            else:
                animals_not_in_foster.add(a_number)

            # Perform these operations last. They will load new pages!
            #
            animal_data[a_number]['sn'] = self._get_spay_neuter_status(a_number)
            animal_data[a_number]['bio'] = 'Yes' if self._animal_has_adoption_summary(a_number) else 'No'

            # Create some helpful/default string representations
            #
            set_default = lambda str, default : default if str.strip() in [None, ''] else str

            animal_data[a_number]['sn'] = set_default(animal_data[a_number]['sn'], 'Unknown')
            animal_data[a_number]['status'] = set_default(animal_data[a_number]['status'], 'Status Unknown')
            animal_data[a_number]['name'] = set_default(animal_data[a_number]['name'], 'Unnamed')
            animal_data[a_number]['breed'] = set_default(animal_data[a_number]['breed'], 'Unknown Breed')
            animal_data[a_number]['gender'] = set_default(animal_data[a_number]['gender'], 'Unknown Gender')
            animal_data[a_number]['color'] = set_default(animal_data[a_number]['primary_color'], 'Unknown Color')

            if animal_data[a_number]['secondary_color'].strip() not in [None, '', 'None']:
                animal_data[a_number]['color'] += '/{}'.format(animal_data[a_number]['secondary_color'])

            breed_abbreviations = {
                # 'Breed' is a free-form text field and not everyone enters data exactly the same. I'll compare against
                # lowercase/no-whitespace for slightly better odds of matches.
                #
                'domesticshorthair'  : 'DSH',
                'domesticmediumhair' : 'DMH',
                'domesticlonghair'   : 'DLH' }

            abbreviation = breed_abbreviations.get(animal_data[a_number]['breed'].replace(' ', '').lower())
            if abbreviation:
                animal_data[a_number]['breed'] = abbreviation

            if animal_data[a_number]['gender'].lower() == 'male':
                animal_data[a_number]['gender_short'] = 'M'
            elif animal_data[a_number]['gender'].lower() == 'female':
                animal_data[a_number]['gender_short'] = 'F'
            else:
                animal_data[a_number]['gender_short'] = animal_data[a_number]['gender']

            if not silent:
                print('{}'.format(animal_data[a_number]['status']))

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
                animal_type = self.YOUNG_ANIMAL_TYPE if age.days <= (365 / 2) else self.ADULT_ANIMAL_TYPE

            else:
                weeks = math.floor(age.days / 7)
                age_string = '{} week{}'.format(weeks, 's' if weeks > 1 else '')
                animal_type = self.YOUNG_ANIMAL_TYPE
        except Exception:
            animal_type = self.BASE_ANIMAL_TYPE
            age_string = 'Unknown Age'

        return age_string, animal_type

    def _get_spay_neuter_status(self, animal_number):
        ''' Load spay/neuter status from the medical details page
        '''
        try:
            self._driver.get(self._medical_details_url.format(animal_number))
            return Utils.utf8(self._get_attr_by_xpath('innerText', '/html/body/table[2]/tbody/tr[2]/td/table/tbody/tr[4]/td[4]'))
        except Exception:
            Log.error('Failed to read spay/neuter status for animal {}'.format(animal_number))
            return 'Unknown'

    def _animal_has_adoption_summary(self, animal_number):
        adoption_summary = ''
        try:
            self._driver.get(self._adoption_summary_url.format(animal_number))
            adoption_summary = self._get_text_by_id('adoptSummary').strip()
        except Exception:
            Log.error('Failed to read adoption summary for animal {}'.format(animal_number))
            return False

        return len(adoption_summary) > 10 # minimum of 10 chars, completely arbitrary in case there is some junk in here

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
        full_name += ' ' if full_name else ''
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
            notes += '{}*** {} is a mentor'.format('\r' if notes else '', full_name)

        match_strings = emails.union([full_name, str(person_number)])
        matching_sheets = self.mentor_sheet_reader.find_matching_mentors(match_strings)
        if matching_sheets:
            notes += '{}*** Found {} matching mentor(s): {}'.format('\r' if notes else '', len(matching_sheets), ', '.join([str(s) for s in matching_sheets]))

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

            FUTURE REFACTOR: Consider combining this with _current_animals_fostered()
            FUTURE REFACTOR: Do not return a mystery tuple. Named dictionary keys or object attribute please.
        '''
        page_number = 1
        previous_foster_count = 0
        euthanized_count = 0
        unassisted_death_count = 0

        while True:
            self._driver.get(self._list_all_animals_url.format(page_number, person_number))
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
        if file:
            file.write('{}\r\n'.format(s))

    def _get_from_dict(self, search_dict, search_key):
        result = [pair[search_key] for pair in search_dict if search_key in pair]
        return result[0] if result else None

    def _get_current_mentee_status(self, arg_mentee_status):
        ''' Get current mentees and mentee status for each mentor
        '''
        autoupdate_completed_mentees = 'autoupdate' in arg_mentee_status # mark 'completed' mentors in the spreadsheet
        Log.success('Looking up mentee status (autoupdate_completed_mentees = {})...'.format(autoupdate_completed_mentees))

        completed_mentees = {}
        current_mentees = self.mentor_sheet_reader.get_current_mentees()
        for current in current_mentees:
            current['active_count'] = 0
            print('Checking mentee status for {}... '.format(current['mentor']), end='', flush=True)

            if current['mentees']:
                for mentee in current['mentees']:
                    current_animal_ids = self._current_animals_fostered(mentee['pid'])
                    mentee['current_animals'] = {}

                    animal_data, _, _ = self._get_animal_data(current_animal_ids, True)

                    for current_animal_id in current_animal_ids:
                        mentee['current_animals'][current_animal_id] = animal_data[current_animal_id]

                    if current_animal_ids:
                        current['active_count'] = current['active_count'] + 1
                    else:
                        completed_mentees.setdefault(current['mentor'], []).append(mentee['pid'])

            days_ago = (datetime.now() - current['most_recent']).days if current['most_recent'] else 'N/A'
            print('active mentees = {}, last assigned days ago = {}'.format(current['active_count'], days_ago))

        if autoupdate_completed_mentees:
            Log.success('Auto-updating completed mentees in the mentor spreadsheet...')
            for mentor in completed_mentees:
                self.mentor_sheet_reader.set_completed_mentees(mentor, completed_mentees[mentor])

        return current_mentees

    def _current_animals_fostered(self, person_number):
        ''' Determine the total number of animals this person is currently fostering. Load the list of all animals this
            person is responsible for, page by page until we have no more pages.
        '''
        page_number = 1
        current_animals = []

        while True:
            self._driver.get(self._responsible_for_paged_url.format(page_number, person_number))
            try:
                table = self._driver.find_element_by_xpath('//*[@id="Table4"]/tbody/tr/td[3]/table[2]')
                rows = table.find_elements(By.TAG_NAME, 'tr')
                if len(rows) < 3:
                    break
                for row in rows:
                    cols = row.find_elements(By.TAG_NAME, 'td')
                    if len(cols) == 12:
                        animal_status = cols[2].text.lower()
                        animal_type = cols[5].text.lower()
                        animal_number = int(cols[3].text)
                        target_types = ['cat', 'kitten'] if not self._dog_mode else ['dog', 'puppy']
                        if 'in foster' in animal_status and animal_status != 'unassisted death - in foster' and animal_type in target_types:
                            animal_number = int(cols[3].text)
                            if animal_number not in current_animals: # ignore duplicates
                                current_animals.append(animal_number)
            except NoSuchElementException:
                break

            page_number = page_number + 1

        return current_animals

    def _output_results(self, animal_data, foster_parents, persons_data, animals_not_in_foster, current_mentee_status, csv_filename):
        ''' Output all of our new super amazing results to a csv file
        '''
        Log.success('Writing results to {}...'.format(csv_filename))
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
            for a_number in animals_with_this_person:
                msg = animal_data[a_number]['message']
                if msg:
                    special_message += '{}{}: {}'.format('\r\r' if special_message else '', a_number, msg)

            cell_number = person_data['cell_phone']
            home_number = person_data['home_phone']
            phone = ''
            if len(cell_number) >= 10: # ignore incomplete phone numbers
                phone = '(C) {}'.format(cell_number)
            if len(home_number) >= 10: # ignore incomplete phone numbers
                phone += '{}(H) {}'.format('\r' if phone else '', home_number)

            emails_str = ''
            for email in person_data['emails']:
                emails_str += '{}{}'.format('\r' if emails_str else '', email)

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
            csv_rows[-1].append('"{}"'.format(emails_str))
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
                                                                      Log.GREEN,
                                                                      report_notes.replace('\r', ', '),
                                                                      Log.END))
        with open(csv_filename, 'w') as outfile:
            for row in csv_rows:
                outfile.write(','.join(row))
                outfile.write('\n')

            if not foster_parents:
                outfile.write('*** None of the animals in this report are currently in foster\n')
                Log.warn('None of the animals in this report are currently in foster. Nothing to do!')

            if animals_not_in_foster:
                outfile.write('\n\n\n*** Animals not in foster\n')
                Log.warn('\nAnimals not in foster')
                for a_number in animals_not_in_foster:
                    outfile.write('{} - {}\n'.format(a_number, animal_data[a_number]['status']))
                    print('{} - {}'.format(a_number, animal_data[a_number]['status']))

            if current_mentee_status:
                outfile.write('\n\nMentor,Active Mentees,Last Assigned (days ago)\n')
                for current in current_mentee_status:
                    days_ago = (datetime.now() - current['most_recent']).days if current['most_recent'] else 'N/A'
                    outfile.write('{},{},{}\n'.format(current['mentor'], current['active_count'], days_ago))

    def _get_animal_details_string(self, foster_animals, animal_data):
        ''' Group animals by type, list useful details for each animal
        '''
        animals_by_type = {}
        for a_number in foster_animals:
            animals_by_type.setdefault(animal_data[a_number]['type'], []).append(a_number)

        animal_details = ''
        for animal_type in animals_by_type:
            animals = animals_by_type[animal_type]
            line = '{} {}{} @ {}'.format(len(animals),
                                         animal_type,
                                         's' if len(animals) > 1 else '',
                                         animal_data[animals[0]]['age'])
            for a_number in sorted(animals):
                line += '\r{} ({}), S/N {}'.format(a_number,
                                                   animal_data[a_number]['gender_short'],
                                                   animal_data[a_number]['sn'])
                if not self._dog_mode:
                    line += ', {}, {}, {}'.format(animal_data[a_number]['name'],
                                                  animal_data[a_number]['breed'],
                                                  animal_data[a_number]['color'])

            animal_details += '{}{}'.format('\r' if animal_details else '', line)

        animal_details_brief = ''
        for a_number in sorted(foster_animals):
            animal_details_brief += '{}{} - {}, {}, {}'.format('\r' if animal_details_brief else '',
                                                               a_number,
                                                               animal_data[a_number]['name'],
                                                               animal_data[a_number]['breed'],
                                                               animal_data[a_number]['color'])
        return animal_details, animal_details_brief

    def _get_text_by_id(self, element_id):
        try:
            return self._driver.find_element_by_id(element_id).text
        except Exception:
            return ''

    def _get_attr_by_id(self, element_id):
        try:
            return self._driver.find_element_by_id(element_id).get_attribute('value')
        except Exception:
            return ''

    def _get_property_by_xpath(self, property_name, element_xpath):
        try:
            return self._driver.find_element_by_xpath(element_xpath).get_property(property_name)
        except Exception:
            return ''

    def _get_attr_by_xpath(self, attr, element_xpath):
        try:
            return self._driver.find_element_by_xpath(element_xpath).get_attribute(attr)
        except Exception:
            return ''

    def _get_checked_by_id(self, element_id):
        try:
            attr = self._driver.find_element_by_id(element_id).get_attribute('checked')
            return bool(attr)
        except Exception:
            return False

    def _get_selection_by_id(self, element_id):
        try:
            select_element = Select(self._driver.find_element_by_id(element_id))
            return select_element.first_selected_option.text
        except Exception:
            return ''

if __name__ == "__main__":
    KittenScraper().run()
