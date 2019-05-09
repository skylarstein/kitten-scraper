# 😺 [Kitten Scraper ](https://github.com/skylarstein/kitten-scraper)Change Log

## [2.0.2](https://github.com/skylarstein/kitten-scraper/compare/v2.0.1...v2.0.2) (2019-05-08)

* Cache Google Sheets data on load (performance optimization)

## [2.0.1](https://github.com/skylarstein/kitten-scraper/compare/v2.0.0...v2.0.1) (2019-05-06)

* Wait for lazy-loaded content, update chromedriver

## [2.0.0](https://github.com/skylarstein/kitten-scraper/compare/v1.8.0...v2.0.0) (2019-03-14)

* Move most configuration to the online mentor spreadsheet. Makes it easier to update the config without new deployment
* General cleanup and improvements
* Update dependencies

## [1.8.0](https://github.com/skylarstein/kitten-scraper/compare/v1.7.0...v1.8.0) (2018-12-13)

* Retrieve current animal status, filter out animals not currently in foster
* Cleanup output CSV, add current animal status, remove unnecessary data, sort
* Improved debug output and error handling
* Fixes Python3 support
* Adds missing dependency

## [1.7.0](https://github.com/skylarstein/kitten-scraper/compare/v1.6.0...v1.7.0) (2018-08-05)

* Support up to 4 emails per foster parent (not covering N emails complexity for now)
* Better handling of empty animal age string
* User friendly error if no VPN connection at login

## [1.6.0](https://github.com/skylarstein/kitten-scraper/compare/v1.5.0...v1.6.0) (2018-07-16)

* Include spay/neuter status with each animal number

## [1.5.0](https://github.com/skylarstein/kitten-scraper/compare/v1.4.0...v1.5.0) (2018-07-13)

* Include 'Special Animal Message'

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