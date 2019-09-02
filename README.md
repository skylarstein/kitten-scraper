# A little bit of web scraping and automation for the Feline Foster Mentor Program

![Platform macOS | Linux](https://img.shields.io/badge/Platform-macOS%20|%20Linux-brightgreen.svg)
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

```text
$ python3 --version
```

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
username : your_username
password : your_password
mentors_spreadsheet_key : key
```

### client_secret.json

Google Sheets integration and Google Sheets API platform access will require a 'client_secret.json' file. [Instructions by pygsheets](https://pygsheets.readthedocs.io/en/stable/authorization.html) will assist you in creating this file. Copy 'client_secret.json' to the kitten-scraper directory.

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
```

## Let's Do This

To generate a report, run kitten-scraper.py from the command line. Specify the path to the original feline foster report xls (-i) as well as the desired output file name (-o).

```text
$ python3 kitten_scraper.py --input ~/Downloads/FosterReport-May12.xls --output ~/Desktop/UpdatedFosterReport-May12.csv
```
