# 😺 [Kitten Scraper ](https://github.com/skylarstein/kitten-scraper)Change Log

## [1.4.0](https://github.com/skylarstein/kitten-scraper/compare/v1.3.0...v1.4.0) (2018-07-11)

* Dismiss alert dialog on Animal Details page if it exists

## [1.3.0](https://github.com/skylarstein/kitten-scraper/compare/v1.2.0...v1.3.0) (2018-06-13)

* Add a note if foster parent is a mentor
* Handle issues with daily reports which occasionally omit data (missing animal type, other corner cases)
* Improved user-friendly error handling (empty daily reports, etc)
* Assure output path exists
* Chromedriver 2.40, Linux support
* General cleanup, organization

## [1.2.0](https://github.com/skylarstein/kitten-scraper/compare/v1.1.0...v1.2.0) (2018-06-03)

* Validate current foster parent ID for each animal ID (daily report is inconsistent)

## [1.1.0](https://github.com/skylarstein/kitten-scraper/compare/v1.0.0...v1.1.0) (2018-05-28)

* Google Sheets integration to associate foster parents with existing mentors
* Add animal numbers to the animal quantity string

## 1.0.0 (2018-05-26)

* Finally got around to tagging a release, yay!
* Handle unknown animal age (sometimes missing from daily reports)
* Add support for do_not_assign_mentor config