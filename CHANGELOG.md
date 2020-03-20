# 😺 [Kitten Scraper](https://github.com/skylarstein/kitten-scraper) Change Log

## [2.1.0](https://github.com/skylarstein/kitten-scraper/compare/v2.0.8...v2.1.0) (2020-03-20)

* Adds Box support for mentors spreadsheet
* Renames "canine mode" to "dog mode" since I like the sound of that better
* Adds "dog mode" support for Mentee Status Report
* Adds animal name to animal status strings
* Adds optional -c/--config command line arg to specify a config file (defaults to 'config.yaml')
* Updates dependencies
* Various small details and improvements

## [2.0.8](https://github.com/skylarstein/kitten-scraper/compare/v2.0.7...v2.0.8) (2019-10-08)

* Adds Windows support

## [2.0.7](https://github.com/skylarstein/kitten-scraper/compare/v2.0.6...v2.0.7) (2019-09-19)

* Default canine_mode = False if not found in config.yaml, allow enable from command line
* Updated chromedriver

## [2.0.6](https://github.com/skylarstein/kitten-scraper/compare/v2.0.5...v2.0.6) (2019-09-09)

* Kitten Scraper now supports canine mode, how about that?

## [2.0.5](https://github.com/skylarstein/kitten-scraper/compare/v2.0.4...v2.0.5) (2019-09-02)

* Mentee status report improvements (remove duplicates, include S/N status)

## [2.0.4](https://github.com/skylarstein/kitten-scraper/compare/v2.0.3...v2.0.4) (2019-08-04)

* Fix under-reported "previous animals fostered" counts
* Add "Current mentee status" function
* It's been fun Python 2, but Python 3 is now required

## [2.0.3](https://github.com/skylarstein/kitten-scraper/compare/v2.0.2...v2.0.3) (2019-07-07)

* Calculate loss rate
* Update to chromedriver 75.0.3770.90, maybe we'll have some luck with the lost-login-cookie issue when running headless

## [2.0.2](https://github.com/skylarstein/kitten-scraper/compare/v2.0.1...v2.0.2) (2019-05-08)

* Cache Google Sheets data on load (performance optimization)

## [2.0.1](https://github.com/skylarstein/kitten-scraper/compare/v2.0.0...v2.0.1) (2019-05-06)

* Wait for lazy-loaded content, update chromedriver to 74.0.3729.6

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
* Chromedriver 2.40.565386, Linux support
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
