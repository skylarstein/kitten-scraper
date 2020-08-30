# A little bit of web scraping and automation for the Feline / Canine Foster Mentor Program

![Platform macOS | Linux | Windows](https://img.shields.io/badge/Platform-macOS%20|%20Linux%20|%20Windows-brightgreen.svg)
![Python | 3.6.x](https://img.shields.io/badge/Python-3.6.x-brightgreen.svg)
![Kitten Machine | Active](https://img.shields.io/badge/Kitten%20Machine-Active-brightgreen.svg)

Kitten-scraper will import the daily Feline Foster report, automatically retrieve additional information for each animal and foster parent, and match foster parents to their existing Feline Foster Mentors.

```text
                                                                     _                ___       _.--.
  _  ___ _   _               ____                                    \`.|\..----...-'`   `-._.-'_.-'`
 | |/ (_) |_| |_ ___ _ __   / ___|  ___ _ __ __ _ _ __   ___ _ __    /  ' `         ,       __.-'
 | ' /| | __| __/ _ \ '_ \  \___ \ / __| '__/ _` | '_ \ / _ \ '__|   )/' _/     \   `-_,   /
 | . \| | |_| ||  __/ | | |  ___) | (__| | | (_| | |_) |  __/ |      `-'" `"\_  ,_.-;_.-\_ ',
 |_|\_\_|\__|\__\___|_| |_| |____/ \___|_|  \__,_| .__/ \___|_|          _.-'_./   {_.'   ; /
                                                 |_|                    {_.-``-'         {_/
```

## Assure you have Python 3 installed

*Note: Depending on your installation, the 'python3' command may be available as 'python'. Same story with the 'pip3' command mentioned below - it may instead simply be 'pip'. You can check the version from the command line as seen below. Use whichever command responds with Python version 3.x.x*

```text
$ python --version
$ python3 --version
```

## Assure you have the Chrome web browser installed

https://www.google.com/chrome/

## Clone the Repository and Install Dependencies

Clone the repository from the command line or optionally [download the zip archive](https://github.com/skylarstein/kitten-scraper/archive/master.zip). Performing 'pip3 install -r requirements.txt' from the command line is required in either case.

```text
$ git clone https://github.com/skylarstein/kitten-scraper.git kitten-scraper
$ cd kitten-scraper
$ pip3 install -r requirements.txt
```

## Setup and Configuration

### config.yaml

Create a text file named 'config.yaml' in the kitten-scraper directory and enter your credentials and configuration in this format:

```yaml
# Required
username : 'your_username'
password : 'your_password'

# Optional. Include only if the mentor spreadsheet lives on Google Sheets.
google_spreadsheet_key : 'key'
google_client_secret : 'client_secret.json'

# Optional. Include only if the mentor spreadsheet lives on Box.
box_user_id : 'id'
box_file_id : 'id'
box_jwt : 'xxx_xxx_config.json'

# Bonus configuration: Dog mode! There are some slight differences when running Kitten Scraper
# with Canine Foster reports. To enable "dog mode", add the following line:
dog_mode : True
```

## Google Sheets Integration

Google Sheets integration and Google Sheets API platform access will require a ```client_secret.json``` file:
1. Sign into your Google account and visit https://developers.google.com/sheets/api/quickstart/python. You only need to follow "Step 1" on this page, as described in the following steps.
2. Click the "Enable the Google Sheets API" button.
3. When asked to "Configure your OAuth client", select "Desktop app".
4. You should then be presented with a "Download Client Configuration" button. Click and save the file into the kitten-scraper folder. The name of the file is not important, but it does need to match the ```google_client_secret``` entry in your config.yaml file.

The ```google_spreadsheet_key``` value for config.yaml value can be found in the URL of the mentor spreadsheet.

## Box Sheets Integration

Box integration requires the creation of a new app via the Box Dev Console.

1. Assure you have 2-step verification enabled for your Box account. [Log in to your account](https://app.box.com/account), scroll down to Authentication, and enable if necessary.

2. Visit the [Dev Console](https://app.box.com/developers/console) and create a new *Enterprise Integration* app. Select "*OAuth 2.0 with JWT (Server Authentication)*", enter any name you like (e.g. Kitten-Scraper or Puppy-Scraper), *Create App*, then *View Your App*.

3. Visit your App's configuration screen
    * From *Application Access*, select *Enterprise*
    * From *Advanced Features*, select *Perform Actions as Users*
    * From *Add and Manage Public Keys*, click *Generate a Public/Private Keypair*. This will download an ```xxx_xxx_config.json``` file. Copy this to your kitten-scraper directory.

4. Vist your App's *General* screen. From the *App Authorization* section, click *Submit for Authorization*. This will generate an email with an API Key. Follow the instructions in the email.

5. The ```box_user_id``` for config.yaml can be found in on your *Account Settings* page, as *Account ID*.

6. The ```box_file_id``` for config.yaml value can be found in the URL of the mentor spreadsheet.

## Command Line Arguments

```text
$ python3 kitten_scraper.py --help
usage: kitten_scraper.py [-h] [-c CONFIG] [-i INPUT] [-s STATUS] [-b]

optional arguments:
  -h, --help            show this help message and exit

  -c CONFIG, --config CONFIG
                        specify a config file (optional, defaults to 'config.yaml')

  -i, --input INPUT     specify the daily foster report (xls), or optionally a comma-separated list of animal numbers

  -s, --mentee_status [verbose,autoupdate,export]
                        retrieve current mentee status
                        'verbose' : includes additional animal details
                        'autoupdate' : flags completed mentees in the mentor spreadsheet
                        'export' : exports mentee status to text file

  -b, --show_browser    show the browser window (generally for debugging)
```

## Let's Do This

To generate a report, run kitten-scraper.py from the command line. Specify the path to the original feline foster report xls (--input) as well as the desired output file name (--output):

```text
$ python3 kitten_scraper.py --input ~/Downloads/FosterReport-May12.xls
```

To create a mentee status report, include --mentee_status with 'export' to save a mentee status file to your Desktop

```text
$ python3 kitten_scraper.py --mentee_status export,verbose
```
