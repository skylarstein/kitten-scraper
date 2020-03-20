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
# with Canine Foster reports. To enable "dog mode", add the following line, or optionally
# include -d or --dog_mode from the command line.
dog_mode : True
```

## Google Sheets Integration

Google Sheets integration and Google Sheets API platform access will require a ```client_secret.json``` file. [Instructions by pygsheets](https://pygsheets.readthedocs.io/en/stable/authorization.html) will assist you in creating this file. Copy ```client_secret.json``` to the kitten-scraper directory.

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
usage: kitten-scraper.py [-h] [-i INPUT] [-o OUTPUT] [--show_browser]

optional arguments:
  -h, --help            show this help message and exit

  -i INPUT, --input INPUT
                        specify the daily kitten report (xls)

  -o OUTPUT, --output OUTPUT
                        specify an output file (csv)

  -s STATUS, --status STATUS
                        save current mentee status to the given file (txt)

  -b, --show_browser    show the browser window (generally for debugging)

  -d, --dog_mode        enable dog mode
```

## Let's Do This

To generate a report, run kitten-scraper.py from the command line. Specify the path to the original feline foster report xls (--input) as well as the desired output file name (--output):

```text
$ python3 kitten_scraper.py --input ~/Downloads/FosterReport-May12.xls --output ~/Desktop/UpdatedFosterReport-May12.csv
```

To create a mentee status report, use the --status option to save status to a given file:

```text
$ python3 kitten_scraper.py --status ~/Desktop/status.txt
```
