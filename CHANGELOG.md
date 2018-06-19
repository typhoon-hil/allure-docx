# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [0.2.0] - 2018-06-13
### Added
- Command line options to control level of details in the generated docx report (issue #14).

## [0.1.3] - 2018-06-13
### Fixed
- Gracefjully treating cases where no test results are found (avoids `ZeroDivisionError`, Issue #15).

## [0.1.2] - 2018-05-30
### Fixed
- Fixed problem when tests doesn't have `start` or `stop` attributes.
  - Also selecting newest test based on json file modification date.

## [0.1.1] - 2018-05-29
### Changed
- Using `shutil.which` to find `OfficeToPDF` or `LibreOfficeToPDF`, as this function will also return batch files.

