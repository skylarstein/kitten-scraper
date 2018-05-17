# Welcome to the Kitten Scraper

![](https://img.shields.io/badge/platform-macOS-brightgreen.svg)
![](https://img.shields.io/badge/Python-2.7.x,%203.6.x-brightgreen.svg)

## Clone the Repository and Install Dependencies

```
% git clone https://github.com/skylarstein/kitten-scraper.git kitten-scraper
% cd kitten-scraper
% pip install -r requirements.txt
```

## Create the Configuration File
Create a text file named 'config.yaml' in the kitten-scraper directory and enter your credentials in this format:

```
username : your_username
password : your_password
```

## Command Line Arguments

```
% python kitten-scraper.py --help
usage: kitten-scraper.py [-h] [-i INPUT] [-o OUTPUT] [--headless]

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