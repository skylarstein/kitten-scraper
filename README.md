# A little bit of web scraping and automation for the Feline / Canine Foster Mentor Program

![Platform macOS | Linux | Windows](https://img.shields.io/badge/Platform-macOS%20|%20Linux%20|%20Windows-brightgreen.svg)
![Python | 3.6.x](https://img.shields.io/badge/Python-3.6.x-brightgreen.svg)
![Kitten Machine | Active](https://img.shields.io/badge/Kitten%20Machine-Active-brightgreen.svg)

Kitten-Scraper will import the daily Feline Foster report, automatically retrieve additional status for each foster animal, determine the current foster parent, and match foster parents to their existing Feline Foster mentors. Bonus feature: "Dog Mode" mode is supported for the Canine Foster program as well.

Why is this program named "Kitten-Scraper"? Good question! This program looks up animal status by automating the Chrome web browser. This is much (*much!*) faster than manually using a web browser yourself... entering your request for each foster animal, each foster parent, pressing enter, copying/pasting the result, etc. Automating a web browser to read information from a web page is call "web scraping". So there we have "Kitten-Scraper".

*Obligatory disclaimer for anyone reading this code: this "quick weekend project" has significantly grown over time. It's time for some refactoring, code re-use, optimizations, and Python modernization. It's also time to start considering some unit tests. That day is not today. I have baby kittens to feed.*

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

*Note: Depending on your installation, the 'python3' command may be available as 'python'. Same situation with the 'pip3' command - it may instead simply be 'pip'. You can check the version from the command line as seen below. Use whichever command responds with Python version 3.x.x*

```text
$ python --version
$ python3 --version
$ pip --version
$ pip3 --version
```

## Assure you have the Chrome web browser installed

https://www.google.com/chrome/

## Assure you have a matching ```chromedriver``` installed

If your Chrome web browser auto-updates or you manually install a new version, you may need to download the latest matching ```chromedriver``` binary and place it in the ```kitten-scraper/bin``` sub-directory that matches your operating system (macOS, Windows, or Linux). The ```chromedriver``` binary allows Kitten-Scraper to automate the Chrome web browser.

https://chromedriver.chromium.org/downloads

## Clone the Repository and Install Dependencies

Clone the repository from the command line or optionally [download the zip archive](https://github.com/skylarstein/kitten-scraper/archive/master.zip). Performing ```pip install -r requirements.txt``` from the command line is required in either case.

```text
$ git clone https://github.com/skylarstein/kitten-scraper.git kitten-scraper
$ cd kitten-scraper
$ pip install -r requirements.txt
```

## Setup and Configuration

### config.yaml

Create a text file named 'config.yaml' in the kitten-scraper directory and enter your credentials and configuration in the format as show below. You can also copy/paste this text as a starting point:

```yaml
# Required
username : 'your_username'
password : 'your_password'

# Required
google_spreadsheet_key : 'key'
google_client_secret : 'client_secret.json'

# Bonus configuration: Dog mode! There are some slight differences when running Kitten-Scraper
# with Canine Foster reports. To enable "dog mode", add the following line.
# For Feline mode, delete this line or set to False
dog_mode : True
```

## Google Sheets Integration

*Warning/Apology: These instructions are likely deprecated. Google really likes to make this part difficult.*

Google Sheets integration and Google Sheets API platform access will require a ```client_secret.json``` file:
1. Sign into your Google account and visit https://developers.google.com/sheets/api/quickstart/python. You only need to follow "Step 1" on this page, as described in the following steps.
2. Click the "Enable the Google Sheets API" button.
3. When asked to "Configure your OAuth client", select "Desktop app".
4. You should then be presented with a "Download Client Configuration" button. Click and save the file into the kitten-scraper folder. The name of the file is not important, but it does need to match the ```google_client_secret``` entry in your config.yaml file.

The ```google_spreadsheet_key``` value for config.yaml value can be found in the URL of the mentor spreadsheet.

## Command Line Arguments

```text
$ python3 kitten_scraper.py --help
usage: kitten_scraper.py [-h] [-c CONFIG] [-i INPUT] [-s STATUS] [-b]

optional arguments:
  -h, --help            show this help message and exit

  -c CONFIG, --config CONFIG
                        specify a config file (optional, defaults to 'config.yaml')

  -i, --input INPUT     specify the daily foster report (xls), or optionally a comma-separated list of animal numbers

  -s, --status [verbose,autoupdate,export]
                        retrieve current mentee status
                        'verbose' : (optional) includes additional animal details
                        'autoupdate' : (optional) marks completed mentees in the mentor spreadsheet
                        'export' : (optional) exports mentee status to text file

  -b, --show_browser    show the web browser window (generally used for debugging)
```

## Let's Do This

To generate a report, run ```kitten-scraper.py``` from the command line. Specify the path to the daily "animals to foster" report xls with ```--input```:

```text
$ python kitten_scraper.py --input ~/Downloads/FosterReport-May12.xls
```

The following ```--status``` command line arguments are optional, and may be combined with or without ```--input```: 

Basic mentee status includes active mentee count and surgery status (if available):

```text
$ python kitten_scraper.py --status
```

Verbose animal status additionally includes bio status, photo status, and age. Note that this will add additional runtime to Kitten-Scraper:

```text
$ python kitten_scraper.py --status "verbose"
```

To auto-update/auto-complete mentees who no longer currently have foster animals, include the "autoupdate" status directive:

```text
$ python kitten_scraper.py --status "autoupdate"
```

Status is always logged to the command line, but to export status to a file on your desktop, include the "export" status directive:

```text
$ python kitten_scraper.py --status "export"
```

Status directives may be combined as comma separated list of commands:

```text
$ python3 kitten_scraper.py --status "verbose,autoupdate,export"
```
