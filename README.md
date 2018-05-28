# Welcome to the Kitten Scraper

![](https://img.shields.io/badge/platform-macOS-brightgreen.svg)
![](https://img.shields.io/badge/Python-2.7.x,%203.6.x-brightgreen.svg)

## Clone the Repository and Install Dependencies

Optionally (instead of git clone), you can [download the zip archive](https://github.com/skylarstein/kitten-scraper/archive/master.zip)

```
% git clone https://github.com/skylarstein/kitten-scraper.git kitten-scraper
% cd kitten-scraper
% pip install -r requirements.txt
```
Don't have pip installed? Two options:

```
$ sudo easy_install pip
```
or..
```
$ curl https://bootstrap.pypa.io/get-pip.py | sudo python
```


## Setup and Configuration

### config.yaml
Create a text file named 'config.yaml' in the kitten-scraper directory and enter your credentials in this format:

```
username : your_username
password : your_password
mentors_spreadsheet_key : key
login_url : http://url
search_url : http://url
list_animals_url : http://url
do_not_assign_mentor : 
    - 100 # special person number
    - 200 # special person number
```
### client_secret.json

For Google Sheets integration and Google Sheets API platform access, copy your client_secret.json file to the kitten-scraper directory.

## Command Line Arguments

```
% python kitten-scraper.py --help
usage: kitten-scraper.py [-h] [-i INPUT] [-o OUTPUT] [--show_browser]

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        daily kitten report (xls)
  -o OUTPUT, --output OUTPUT
                        output file (csv)
  --show_browser        show the browser window while working
```
## Let's Do This

```
% python kitten-scraper.py -i /Users/skylar/Downloads/FosterReport.xls -o FosterReport.csv