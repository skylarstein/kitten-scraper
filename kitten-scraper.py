import os
import sys
import time
import xlrd
import yaml
from datetime import datetime
from argparse import ArgumentParser
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

class KittenScraper():

	def __init__(self):
		self.LOGIN_URL = 'http://192.168.100.27/login.aspx?aspInitiated=1.1'
		self.SEARCH_URL = 'http://192.168.100.27/main.asp'
		self.CONFIG_FILE = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'config.yaml')

	def close(self):
		self.driver.close()
		self.driver.quit()

	def load_configuration(self):
		try:
			config = yaml.load(file(self.CONFIG_FILE, 'r'))
			self.username = config['username']
			self.password = config['password']
			return True

		except yaml.YAMLError as err:
			print 'ERROR: Unable to parse configuration file: {}, {}'.format(self.CONFIG_FILE, err)
		
		except IOError as err:
			print 'ERROR: Unable to read configuration file: {}, {}'.format(self.CONFIG_FILE, err)

		return False

	def start_browser(self, headless):
		chrome_options = webdriver.ChromeOptions()
		chrome_options.add_argument('--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.117 Safari/537.36')
		if headless:
			chrome_options.add_argument("--headless")

		chromedriver_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'bin/macosx/chromedriver')
		self.driver = webdriver.Chrome(chromedriver_path, chrome_options = chrome_options)

	def login(self):
		print 'Logging in...'
		self.driver.get(self.LOGIN_URL)

		self.driver.find_element_by_id("txt_username").send_keys(self.username)
		self.driver.find_element_by_id("txt_password").send_keys(self.password)
		self.driver.find_element_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_btn_login').click()
		self.driver.find_element_by_id('Continue').click()

	def get_person_data(self, person_number):
		print 'Looking up person number {}...'.format(person_number)

		self.driver.get(self.SEARCH_URL)
		self.driver.find_element_by_id("userid").send_keys(person_number)
		self.driver.find_element_by_id("userid").send_keys(Keys.RETURN)

		first_name      = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtFirstName')
		last_name       = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtLastName')
		preferred_name  = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonNameTitle1_txtPreferredName')
		home_phone      = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_homePhone_txtPhone3')
		cell_phone      = self.get_text_by_id('ctl00_ctl00_ContentPlaceHolderBase_ContentPlaceHolder1_personDetailsUC_PersonContact1_mobilePhone_txtPhone3')		
		primary_email   = self.get_text_by_xpath('//*[@id="emailTable"]/tbody/tr[1]/td[1]')
		secondary_email = self.get_text_by_xpath('//*[@id="emailTable"]/tbody/tr[2]/td[1]')

		return {
			'first_name'      : first_name,
			'last_name'       : last_name,
			'preferred_name'  : preferred_name,
			'home_phone'      : home_phone,
			'cell_phone'      : cell_phone,
			'primary_email'   : primary_email,
			'secondary_email' : secondary_email
		}

	def get_text_by_id(self, element_id):
		try:
			return self.driver.find_element_by_id(element_id).get_attribute('value')
		except:
			return ''

	def get_text_by_xpath(self, element_xpath):
		try:
			return self.driver.find_element_by_xpath(element_xpath).get_attribute('innerText')
		except:
			return ''

class KittenReportReader:

	def load_xls(self, xls_filename):
		try:
			self.workbook = xlrd.open_workbook(xls_filename)
			self.sheet = self.workbook.sheet_by_index(0)

			if (self.sheet.row_values(0)[0] != 'Datetime of Current Status Date' or
			    self.sheet.row_values(0)[1] != 'Current Animal Type' or
			    self.sheet.row_values(0)[2] != 'AnimalID' or
			    self.sheet.row_values(0)[3] != 'Animal Name' or
			    self.sheet.row_values(0)[4] != 'Age' or
			    self.sheet.row_values(0)[5] != 'Foster Parent ID'):

				print 'ERROR: Unexpected column layout in {}'.format(xls_filename)
				return False

			print 'Loaded report {}'.format(xls_filename)
			return True

		except IOError as err:
			print 'ERROR: Unable to read xls file: {}, {}'.format(xls_filename, err)

		except xlrd.XLRDError as err:
			print 'ERROR: Unable to read xls file: {}'.format(err.message)

		return False

	def get_person_numbers(self):
		persons = set()
		for n in range(1, self.sheet.nrows):
			# If a person number has no associated animal number, this is due to a bug in the
			# report which includes a mostly empty row for an animal's previous foster parent
			#
			animal_number = self.sheet.row_values(n)[2]
			person_number = self.sheet.row_values(n)[5]

			if isinstance(animal_number, float): # xls stores all numbers as float
				persons.add(str(int(person_number)))

		return persons

	def copy_row_as_text(self, row_number):
		''' Output is going to CSV so we need to stringify all types (dates in particular)
		'''
		values = []

		for col_number in range(0, len(self.sheet.row_values(row_number))):
			cell_type = self.sheet.cell_type(row_number, col_number)

			if cell_type == xlrd.XL_CELL_DATE:
				dt = datetime(*xlrd.xldate_as_tuple(self.sheet.row_values(row_number)[col_number], self.workbook.datemode))
				# wrapping datestr in ="%s" to stop Excel from trying to be smart with date strings
				values.append(dt.strftime('="%d-%b-%Y %-I:%M %p"'))

			elif cell_type == xlrd.XL_CELL_NUMBER:
				values.append(str(int(self.sheet.row_values(row_number)[col_number])))

			else:
				s = str(self.sheet.row_values(row_number)[col_number])
				if s == "null":
					s = ''
				values.append('"{}"'.format(s))

		return values

	def output_results(self, persons_data, csv_filename):
		print 'Writing results to {}...'.format(csv_filename)

		new_rows = []
	
		# Start with the original column headers and then add our new ones
		#
		new_rows.append(self.sheet.row_values(0))
		new_rows[-1].append('Name')
		new_rows[-1].append('E-mail')
		new_rows[-1].append('Phone')
		new_rows[-1].append('Foster Experience')
		new_rows[-1].append('Date Kittens Received')	
		new_rows[-1].append('Quantity')

		for row_number in range(1, self.sheet.nrows):
			# If there is no animal number in this row, skip the row
			#
			if not self.sheet.row_values(row_number)[2]:
				continue

			# Include original column data as text since we're building a CSV document
			#
			new_rows.append(self.copy_row_as_text(row_number))

			# Only include person details for rows with 'Current Animal Type' populated
			#
			if not self.sheet.row_values(row_number)[1]:
				continue

			# Grab the person data from their person number
			#
			person_number = str(int(self.sheet.row_values(row_number)[5]))
			person_data = persons_data[person_number] if person_number in persons_data else {}

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
			if len(cell_number):
				phone = 'c: {}'.format(cell_number)

			if len(home_number):
				if len(phone):
					phone += '\r\n'
				phone += 'h: {}'.format(home_number)

			# Build email(s)
			#
			email = person_data['primary_email'] if 'primary_email' in person_data else ''
			secondary_email = person_data['secondary_email'] if 'secondary_email' in person_data else ''

			if len(secondary_email):
				if len(email):
					email += '\r\n'
				email += secondary_email

			new_rows[-1].append('"{}"'.format(name))
			new_rows[-1].append('"{}"'.format(email))
			new_rows[-1].append('"{}"'.format(phone))
			new_rows[-1].append('') # TODO foster experience
			new_rows[-1].append('') # TODO date kittens received
			new_rows[-1].append('') # TODO quantity

		with open(csv_filename, 'w') as outfile:
			for row in new_rows:
				outfile.write(','.join(row))
				outfile.write('\n')

if __name__ == "__main__":
	start_time = time.time()

	arg_parser = ArgumentParser()
	arg_parser.add_argument('-i', '--input', help = 'Daily kitten report (xls)', required = False)
	arg_parser.add_argument('-o', '--output', help = 'Output file (csv)', required = False)
	arg_parser.add_argument('--headless', help = 'Run headless browser', required = False, action = 'store_true')
	args = arg_parser.parse_args()

	if not args.input or not args.output:
		arg_parser.print_help()
		sys.exit(0)

	# Load me up some kittens and foster parent numbers from the XLS report
	#
	kitten_reader = KittenReportReader()
	if not kitten_reader.load_xls(args.input):
		sys.exit()

	persons = kitten_reader.get_person_numbers()
	print 'Found foster parent numbers: {}'.format(persons)

	# Log in and query foster parent details (person number -> name and contact details)
	#
	kitten_scraper = KittenScraper()
	if not kitten_scraper.load_configuration():
		sys.exit()

	kitten_scraper.start_browser(args.headless)
	kitten_scraper.login()

	persons_data = {}
	for person in persons:
		persons_data[person] = kitten_scraper.get_person_data(person)

	kitten_scraper.close()
	
	# Output the combined results to CSV
	#
	kitten_reader.output_results(persons_data, args.output)

	print '\nKitten scraping completed in {0:.3f} seconds'.format(time.time() - start_time)
